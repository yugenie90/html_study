/* global ExcelJS */

/**
 * ✅ 전제(사용자 편집):
 * - 강의정보 시트에서 목동/기숙 강의실 행은 삭제하고 업로드
 *
 * ✅ 처리:
 * - 학생정보: 헤더 기반(자동 탐지) + 대치 관만 처리
 * - 강의정보: 헤더 기반(자동 탐지) + 유형=선택 제외
 * - 필수(반 기준 고정) 시간표 블로킹(요일+교시 충돌 금지)
 * - 탐구: 조합(반+관+탐구1+탐구2) 그룹 우선 배정
 * - 정원 초과 시 분할 배정 + 경고(빨간 행)
 */

const $file = document.getElementById("file");
const $run = document.getElementById("run");
const $download = document.getElementById("download");
const $log = document.getElementById("log");

let inputArrayBuffer = null;
let outputArrayBuffer = null;

const DAECHI_BUILDINGS = new Set(["W관", "N관", "S관", "브릿지관", "M3관", "3H관", "신관"]);

function log(msg) { $log.textContent += msg + "\n"; }
function norm(v) { return (v ?? "").toString().trim(); }
function cellText(cell) { return norm(cell?.text ?? cell?.value ?? ""); }

const dayOrder = new Map([
  ["월요일", 1], ["화요일", 2], ["수요일", 3], ["목요일", 4], ["금요일", 5], ["토요일", 6], ["일요일", 7],
]);

function periodNum(p) {
  const m = norm(p).match(/(\d+)/);
  return m ? parseInt(m[1], 10) : 99;
}

async function readWorkbook(arrayBuffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(arrayBuffer);
  return wb;
}

/* --------- 헤더 자동 탐지 + 시트 읽기 --------- */
function findHeaderRowNumber(ws, requiredCols, maxScanRows = 12) {
  let bestRow = 1, bestScore = -1;
  for (let r = 1; r <= Math.min(maxScanRows, ws.rowCount); r++) {
    const row = ws.getRow(r);
    const headers = [];
    row.eachCell((cell, col) => { headers[col] = cellText(cell); });
    const set = new Set(headers.filter(Boolean));
    const score = requiredCols.reduce((acc, c) => acc + (set.has(c) ? 1 : 0), 0);
    if (score > bestScore) { bestScore = score; bestRow = r; }
  }
  return bestRow;
}

function sheetToObjects(ws, requiredCols) {
  const headerRowNumber = findHeaderRowNumber(ws, requiredCols, 12);
  const headerRow = ws.getRow(headerRowNumber);

  const headers = [];
  headerRow.eachCell((cell, col) => { headers[col] = cellText(cell); });

  const rows = [];
  for (let r = headerRowNumber + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const anyVal = row.values?.some(v => v !== null && v !== undefined && norm(v) !== "");
    if (!anyVal) continue;

    const obj = {};
    let empty = true;
    headers.forEach((h, col) => {
      if (!h) return;
      const val = cellText(row.getCell(col));
      obj[h] = val;
      if (val !== "") empty = false;
    });
    if (!empty) rows.push(obj);
  }
  return rows;
}

/* --------- 시간표/정원 유틸 --------- */
function slotsFromRows(rows) {
  const set = new Set();
  for (const r of rows) set.add(`${norm(r["요일"])}|${norm(r["교시"])}`);
  return set;
}

function hasConflict(slotsA, slotsB) {
  for (const s of slotsA) if (slotsB.has(s)) return true;
  return false;
}

function buildDetail(rows) {
  const sorted = [...rows].sort((a,b) => {
    const da = dayOrder.get(norm(a["요일"])) ?? 99;
    const db = dayOrder.get(norm(b["요일"])) ?? 99;
    if (da !== db) return da - db;
    return periodNum(a["교시"]) - periodNum(b["교시"]);
  });
  return sorted.map(r => `${norm(r["요일"])}${norm(r["교시"])}(${norm(r["강의실"])})`).join(" / ");
}

function minCapacity(rows) {
  let cap = null;
  for (const r of rows) {
    const n = Number(norm(r["최대인원"]));
    if (!Number.isFinite(n)) continue;
    cap = (cap === null) ? n : Math.min(cap, n);
  }
  return cap;
}

/* --------- 메인 배정 --------- */
async function runAssignment(inputWb) {
  const lecWs = inputWb.getWorksheet("강의정보");
  const stuWs = inputWb.getWorksheet("학생정보");
  if (!lecWs || !stuWs) throw new Error("시트 이름이 정확해야 합니다: 강의정보, 학생정보");

  const lecRequired = ["과목","유형","반","강의실","요일","교시","최대인원"];
  const stuRequired = ["일련번호","관","반","탐구1","탐구2"];

  const lecAll = sheetToObjects(lecWs, lecRequired);
  const stuAll = sheetToObjects(stuWs, stuRequired);

  log(`강의정보 ${lecAll.length}행, 학생정보 ${stuAll.length}행 로드`);

  const missingL = lecRequired.filter(c => !Object.prototype.hasOwnProperty.call(lecAll[0] ?? {}, c));
  if (missingL.length) throw new Error(`강의정보 컬럼 누락: ${missingL.join(", ")}`);
  const missingS = stuRequired.filter(c => !Object.prototype.hasOwnProperty.call(stuAll[0] ?? {}, c));
  if (missingS.length) throw new Error(`학생정보 컬럼 누락: ${missingS.join(", ")}`);

  // 선택 제외
  const lec = lecAll.filter(r => norm(r["유형"]) !== "선택");

  // 대치 관 학생만 처리
  const stu = stuAll.filter(s => DAECHI_BUILDINGS.has(norm(s["관"])));
  log(`대치 관 학생만 처리: ${stu.length}/${stuAll.length}`);

  // 1) 필수 시간표(반 기준 고정)
  const mandatoryRows = lec.filter(r => norm(r["유형"]) === "필수");
  const mandatoryByHome = new Map();
  for (const r of mandatoryRows) {
    const home = norm(r["반"]);
    if (!mandatoryByHome.has(home)) mandatoryByHome.set(home, []);
    mandatoryByHome.get(home).push(r);
  }

  const mandatorySlotsByHome = new Map();
  const mandatoryDetailByHome = new Map();
  for (const [home, rows] of mandatoryByHome.entries()) {
    mandatorySlotsByHome.set(home, slotsFromRows(rows));
    mandatoryDetailByHome.set(home, rows
      .sort((a,b)=> (dayOrder.get(norm(a["요일"]))??99)-(dayOrder.get(norm(b["요일"]))??99) || periodNum(a["교시"])-periodNum(b["교시"]))
      .map(r => `${norm(r["과목"])} ${norm(r["요일"])}${norm(r["교시"])}(${norm(r["강의실"])})`)
      .join(" / "));
  }

  // 2) 탐구 반 정보
  const exploreRows = lec.filter(r => norm(r["유형"]) === "탐구");
  const exploreBySection = new Map();
  for (const r of exploreRows) {
    const sec = norm(r["반"]);
    if (!exploreBySection.has(sec)) exploreBySection.set(sec, []);
    exploreBySection.get(sec).push(r);
  }

  const sections = new Map();
  for (const [sec, rows] of exploreBySection.entries()) {
    sections.set(sec, {
      sec,
      slots: slotsFromRows(rows),
      detail: buildDetail(rows),
      capacity: minCapacity(rows),
    });
  }

  const counts = new Map([...sections.keys()].map(s => [s, 0]));
  function remainingCap(sec) {
    const cap = sections.get(sec).capacity;
    if (cap === null || cap === undefined || Number.isNaN(Number(cap))) return Infinity;
    return Number(cap) - (counts.get(sec) ?? 0);
  }
  function canFit(sec, n) { return remainingCap(sec) >= n; }

  function getCandidates(subject, blockedSlots, groupSize) {
    const s = norm(subject);
    if (!s) return [];
    const out = [];
    for (const [sec, info] of sections.entries()) {
      if (!norm(sec).startsWith(s)) continue;
      if (!canFit(sec, groupSize)) continue;
      if (hasConflict(info.slots, blockedSlots)) continue;
      out.push(sec);
    }
    out.sort((a,b)=> (counts.get(a)-counts.get(b)) || a.localeCompare(b));
    return out;
  }

  function scorePair(sec1, sec2, groupSize) {
    const c1 = (counts.get(sec1) ?? 0) + groupSize;
    const c2 = (counts.get(sec2) ?? 0) + groupSize;
    return [Math.max(c1,c2), c1+c2, sec1, sec2];
  }

  // 3) 그룹(반+관+탐구1+탐구2)
  const groups = new Map();
  for (const s of stu) {
    const key = `${norm(s["반"])}||${norm(s["관"])}||${norm(s["탐구1"])}||${norm(s["탐구2"])}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(s);
  }

  const groupEntries = [...groups.entries()]
    .map(([key, members]) => ({ key, members }))
    .sort((a,b)=> b.members.length - a.members.length);

  const results = [];
  const groupPairMap = new Map(); // key -> {sec1, sec2}

  function tryPairForSize(home, subj1, subj2, size, preferredPair) {
    const blockedBase = mandatorySlotsByHome.get(home) ?? new Set();

    // preferred 우선
    if (preferredPair) {
      const {sec1, sec2} = preferredPair;
      if (sec1 && sec2 && sections.has(sec1) && sections.has(sec2) && canFit(sec1,size) && canFit(sec2,size)) {
        if (!hasConflict(sections.get(sec1).slots, blockedBase)) {
          const blocked = new Set(blockedBase);
          for (const x of sections.get(sec1).slots) blocked.add(x);
          if (!hasConflict(sections.get(sec2).slots, blocked)) return {sec1, sec2};
        }
      }
    }

    // search best pair
    const cand1 = getCandidates(subj1, blockedBase, size);
    let best = null;
    for (const sec1 of cand1) {
      const blocked = new Set(blockedBase);
      for (const x of sections.get(sec1).slots) blocked.add(x);
      const cand2 = getCandidates(subj2, blocked, size);
      for (const sec2 of cand2) {
        const sc = scorePair(sec1, sec2, size);
        if (!best || sc < best.score) best = {sec1, sec2, score: sc};
      }
    }
    return best ? {sec1: best.sec1, sec2: best.sec2} : null;
  }

  function assignGroup(key, members) {
    const [home, building, subj1, subj2] = key.split("||");
    const groupSize = members.length;

    let warnBase = false;
    const notesBase = [];

    if (!mandatorySlotsByHome.has(home)) {
      warnBase = true;
      notesBase.push(`필수시간표 없음(반=${home})`);
    }

    const preferred = groupPairMap.get(key);

    // 1) 전체 그룹 시도
    const full = tryPairForSize(home, subj1, subj2, groupSize, preferred);
    if (full) {
      counts.set(full.sec1, (counts.get(full.sec1)??0) + groupSize);
      counts.set(full.sec2, (counts.get(full.sec2)??0) + groupSize);
      groupPairMap.set(key, {sec1: full.sec1, sec2: full.sec2});
      return [{ size: groupSize, sec1: full.sec1, sec2: full.sec2, warn: warnBase, notes: notesBase }];
    }

    // 2) 분할(최소 분할) — 큰 chunk부터
    let remaining = groupSize;
    const chunks = [];
    let split = false;

    while (remaining > 0) {
      let found = null;
      for (let size = remaining; size >= 1; size--) {
        const p = tryPairForSize(home, subj1, subj2, size, preferred);
        if (p) { found = { size, sec1: p.sec1, sec2: p.sec2 }; break; }
      }

      if (!found) {
        split = true;
        chunks.push({
          size: remaining,
          sec1: "",
          sec2: "",
          warn: true,
          notes: [...notesBase, "조합 단위 배정 실패(정원/충돌/개설없음)"]
        });
        break;
      }

      if (found.size !== groupSize) split = true;
      counts.set(found.sec1, (counts.get(found.sec1)??0) + found.size);
      counts.set(found.sec2, (counts.get(found.sec2)??0) + found.size);

      chunks.push({
        size: found.size,
        sec1: found.sec1,
        sec2: found.sec2,
        warn: warnBase || split,
        notes: split ? [...notesBase, "⚠️ 같은 조합 그룹이 정원/제약으로 분할 배정됨"] : notesBase
      });

      remaining -= found.size;
    }

    return chunks;
  }

  for (const g of groupEntries) {
    const chunks = assignGroup(g.key, g.members);

    let idx = 0;
    for (const ch of chunks) {
      const part = g.members.slice(idx, idx + ch.size);
      idx += ch.size;

      for (const stuRow of part) {
        const id = norm(stuRow["일련번호"]);
        const home = norm(stuRow["반"]);
        const building = norm(stuRow["관"]);
        const subj1 = norm(stuRow["탐구1"]);
        const subj2 = norm(stuRow["탐구2"]);

        const warn = ch.warn || !ch.sec1 || !ch.sec2;
        const notes = [...(ch.notes ?? [])];
        if (!ch.sec1) notes.push(`탐구1(${subj1}) 미배정`);
        if (!ch.sec2) notes.push(`탐구2(${subj2}) 미배정`);

        results.push({
          "일련번호": id,
          "관": building,
          "학생반": home,
          "탐구1": subj1,
          "탐구1_배정반": ch.sec1,
          "탐구1_시간표": ch.sec1 ? sections.get(ch.sec1).detail : "",
          "탐구2": subj2,
          "탐구2_배정반": ch.sec2,
          "탐구2_시간표": ch.sec2 ? sections.get(ch.sec2).detail : "",
          "경고": warn ? "⚠️" : "",
          "비고": notes.filter(Boolean).join(" | "),
          "필수시간표(참고)": mandatoryDetailByHome.get(home) ?? "",
          "조합그룹키(참고)": g.key
        });
      }
    }
  }

  // 반별명단
  const roster = [];
  for (const [sec, info] of sections.entries()) {
    const members = results.filter(r => r["탐구1_배정반"] === sec || r["탐구2_배정반"] === sec);
    if (!members.length) continue;
    for (const m of members) {
      roster.push({
        "탐구반": sec,
        "일련번호": m["일련번호"],
        "학생반": m["학생반"],
        "관": m["관"],
        "탐구1": m["탐구1"],
        "탐구2": m["탐구2"],
        "정원": info.capacity ?? "",
        "탐구반_시간표": info.detail,
      });
    }
  }

  // 요약
  const summary = [];
  for (const [sec, info] of sections.entries()) {
    const c = counts.get(sec) ?? 0;
    if (c <= 0) continue;
    summary.push({
      "탐구반": sec,
      "배정인원": c,
      "정원": info.capacity ?? "",
      "시간표": info.detail,
    });
  }
  summary.sort((a,b)=> b["배정인원"] - a["배정인원"] || norm(a["탐구반"]).localeCompare(norm(b["탐구반"])));

  // 결과 엑셀 생성
  const outWb = new ExcelJS.Workbook();
  outWb.created = new Date();

  function addSheet(name, rows, warnRed = false) {
    const ws = outWb.addWorksheet(name);
    if (!rows || rows.length === 0) { ws.addRow(["데이터 없음"]); return ws; }

    const cols = Object.keys(rows[0]);
    ws.addRow(cols);

    const header = ws.getRow(1);
    header.font = { bold: true, color: { argb: "FFFFFFFF" } };
    header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E79" } };
    header.alignment = { vertical: "middle", horizontal: "center" };

    for (const r of rows) ws.addRow(cols.map(c => r[c]));

    ws.views = [{ state: "frozen", ySplit: 1 }];
    cols.forEach((c, idx) => { ws.getColumn(idx + 1).width = Math.min(Math.max(12, c.length + 2), 60); });

    if (warnRed) {
      const warnIdx = cols.indexOf("경고") + 1;
      if (warnIdx > 0) {
        for (let i = 2; i <= ws.rowCount; i++) {
          const cell = ws.getRow(i).getCell(warnIdx);
          if (norm(cell.value) === "⚠️") {
            ws.getRow(i).eachCell((c) => {
              c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } };
            });
          }
        }
      }
    }
    return ws;
  }

  addSheet("학생별배정", results, true);
  addSheet("반별명단", roster, false);
  addSheet("요약", summary, false);

  return await outWb.xlsx.writeBuffer();
}

/* --------- 다운로드 --------- */
function downloadBuffer(buf, filename) {
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/* --------- UI 이벤트 --------- */
$file.addEventListener("change", async (e) => {
  const f = e.target.files?.[0];
  outputArrayBuffer = null;
  $download.disabled = true;
  $run.disabled = true;
  $log.textContent = "";

  if (!f) return;
  log(`파일 선택됨: ${f.name}`);
  inputArrayBuffer = await f.arrayBuffer();
  $run.disabled = false;
});

$run.addEventListener("click", async () => {
  try {
    $run.disabled = true;
    log("워크북 로딩 중...");
    const wb = await readWorkbook(inputArrayBuffer);

    log("배정 실행 중...");
    outputArrayBuffer = await runAssignment(wb);

    log("완료! 결과 다운로드 버튼을 누르세요.");
    $download.disabled = false;
  } catch (err) {
    console.error(err);
    log(`에러: ${err.message ?? err}`);
    $run.disabled = false;
  }
});

$download.addEventListener("click", () => {
  if (!outputArrayBuffer) return;
  const name = `탐구반_자동배정_결과_${new Date().toISOString().slice(0, 10)}.xlsx`;
  downloadBuffer(outputArrayBuffer, name);
});
