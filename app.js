/* global ExcelJS */

/**
 * ✅ 통짜 완성본(app.js)
 * - 업로드 엑셀 1개 (시트: 강의정보, 학생정보)
 * - 강의정보: 열 위치 고정으로 읽음(헤더 무시)  ✅ L=최대인원, M=소속관, N=지역
 * - 학생정보: 헤더 자동 탐지(1~12행 스캔) 후 읽음
 * - 유형=선택 제외
 * - 필수(반 기준 고정) 시간표 블로킹 후 탐구 배정(충돌 금지)
 * - 조합(반+관+탐구1+탐구2) 그룹 단위로 (탐구1반, 탐구2반) 최대한 동일 유지
 * - 정원 초과 시 가능한 만큼 먼저 배정 → 남는 인원은 다른 조합으로 분할 배정 + 경고(빨간 행)
 * - 지역/소속관 제약 적용(학생 관 → 지역 자동 매핑 + 강의 지역/소속관 매칭)
 */

const $file = document.getElementById("file");
const $run = document.getElementById("run");
const $download = document.getElementById("download");
const $log = document.getElementById("log");

let inputArrayBuffer = null;
let outputArrayBuffer = null;

function log(msg) {
  $log.textContent += msg + "\n";
}

function norm(v) {
  return (v ?? "").toString().trim();
}

function cellText(cell) {
  return norm(cell?.text ?? cell?.value ?? "");
}

function periodNum(p) {
  const m = norm(p).match(/(\d+)/);
  return m ? parseInt(m[1], 10) : 99;
}

const dayOrder = new Map([
  ["월요일", 1], ["화요일", 2], ["수요일", 3], ["목요일", 4], ["금요일", 5], ["토요일", 6], ["일요일", 7],
]);

function regionFromBuilding(building) {
  const b = norm(building);

  // 대치
  const daechi = new Set(["W관", "N관", "S관", "브릿지관", "M3관", "3H관", "신관"]);
  // 목동
  const mokdong = new Set(["목동관", "목동W관"]);
  // 기숙
  const dorm = new Set(["기숙관"]);

  if (daechi.has(b)) return "대치";
  if (mokdong.has(b)) return "목동";
  if (dorm.has(b)) return "기숙";
  return "";
}

async function readWorkbook(arrayBuffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(arrayBuffer);
  return wb;
}

/** =========================
 * 학생정보: 헤더 자동 탐지
 ========================= */
function findHeaderRowNumber(ws, requiredCols, maxScanRows = 12) {
  let bestRow = 1;
  let bestScore = -1;

  for (let r = 1; r <= Math.min(maxScanRows, ws.rowCount); r++) {
    const row = ws.getRow(r);
    const headers = [];
    row.eachCell((cell, col) => { headers[col] = cellText(cell); });

    const headerSet = new Set(headers.filter(Boolean));
    const score = requiredCols.reduce((acc, c) => acc + (headerSet.has(c) ? 1 : 0), 0);

    if (score > bestScore) {
      bestScore = score;
      bestRow = r;
    }
  }
  return bestRow;
}

function sheetToObjectsAutoHeader(ws, requiredColsForDetect) {
  const headerRowNumber = findHeaderRowNumber(ws, requiredColsForDetect, 12);
  const headerRow = ws.getRow(headerRowNumber);

  const headers = [];
  headerRow.eachCell((cell, col) => { headers[col] = cellText(cell); });

  const rows = [];
  for (let r = headerRowNumber + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    // 완전 빈 행 스킵
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

/** =========================
 * 강의정보: 열 위치 고정
 * (스크린샷 기준)
 * A:순번 B:강의코드 C:강의실총등 D:요일&교시&강의실 E:과목 F:유형 G:선생님 H:반 I:강의실 J:요일 K:교시 L:최대인원 M:소속관 N:지역
 ========================= */
function lectureSheetToObjectsFixed(ws, headerRowNumber = 1) {
  const rows = [];

  for (let r = headerRowNumber + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);

    // 완전 빈 행 스킵
    const anyVal = row.values?.some(v => v !== null && v !== undefined && norm(v) !== "");
    if (!anyVal) continue;

    const obj = {
      // 참고로 필요한 것만 최소 구성
      "요일&교시&강의실": cellText(row.getCell(4)),  // D
      "과목": cellText(row.getCell(5)),              // E
      "유형": cellText(row.getCell(6)),              // F
      "선생님": cellText(row.getCell(7)),            // G (필수는 아님)
      "반": cellText(row.getCell(8)),                // H
      "강의실": cellText(row.getCell(9)),            // I (실제 장소)
      "요일": cellText(row.getCell(10)),             // J
      "교시": cellText(row.getCell(11)),             // K
      "최대인원": cellText(row.getCell(12)),         // L ✅
      "소속관": cellText(row.getCell(13)),           // M ✅
      "지역": cellText(row.getCell(14)),             // N ✅
    };

    // 데이터행 유효성(최소한 유형/반/과목 중 하나는 있어야)
    if (!obj["유형"] && !obj["반"] && !obj["과목"]) continue;

    rows.push(obj);
  }

  return rows;
}

/** =========================
 * 공통 유틸
 ========================= */
function slotsFromRows(rows) {
  const set = new Set();
  for (const r of rows) {
    const key = `${norm(r["요일"])}|${norm(r["교시"])}`;
    set.add(key);
  }
  return set;
}

function hasConflict(slotsA, slotsB) {
  for (const s of slotsA) {
    if (slotsB.has(s)) return true;
  }
  return false;
}

function buildDetail(rows) {
  const sorted = [...rows].sort((a, b) => {
    const da = dayOrder.get(norm(a["요일"])) ?? 99;
    const db = dayOrder.get(norm(b["요일"])) ?? 99;
    if (da !== db) return da - db;
    return periodNum(a["교시"]) - periodNum(b["교시"]);
  });
  return sorted
    .map(r => `${norm(r["요일"])}${norm(r["교시"])}(${norm(r["강의실"])})`)
    .join(" / ");
}

function minCapacity(rows) {
  let cap = null;
  for (const r of rows) {
    const n = Number(norm(r["최대인원"]));
    if (!Number.isFinite(n)) continue;
    cap = (cap === null) ? n : Math.min(cap, n);
  }
  return cap; // null 가능
}

async function runAssignment(inputWb) {
  const lecWs = inputWb.getWorksheet("강의정보");
  const stuWs = inputWb.getWorksheet("학생정보");
  if (!lecWs || !stuWs) {
    throw new Error("시트 이름이 정확해야 합니다: 강의정보, 학생정보");
  }

  // ✅ 강의정보: 열 고정 (헤더 무시)
  const lecAll = lectureSheetToObjectsFixed(lecWs, 1);

  // ✅ 학생정보: 헤더 자동 탐지
  const requiredStuCols = ["일련번호", "관", "반", "탐구1", "탐구2"];
  const stuAll = sheetToObjectsAutoHeader(stuWs, requiredStuCols);

  log(`강의정보 ${lecAll.length}행, 학생정보 ${stuAll.length}행 로드`);

  // 학생정보 최소 컬럼 검사(첫 행 기준)
  const missingStu = requiredStuCols.filter(c => !Object.prototype.hasOwnProperty.call(stuAll[0] ?? {}, c));
  if (missingStu.length > 0) {
    throw new Error(`학생정보 컬럼 누락: ${missingStu.join(", ")}`);
  }

  // 강의정보 필수 키 검사(열 고정이라 거의 안 깨짐)
  const requiredLecKeys = ["과목", "유형", "반", "강의실", "요일", "교시", "최대인원", "소속관", "지역"];
  const missingLec = requiredLecKeys.filter(c => !Object.prototype.hasOwnProperty.call(lecAll[0] ?? {}, c));
  if (missingLec.length > 0) {
    throw new Error(`강의정보 열 매핑 실패(열 위치 확인 필요): ${missingLec.join(", ")}`);
  }

  // 0) 유형=선택 제외
  const lec = lecAll.filter(r => norm(r["유형"]) !== "선택");

  // 1) 필수 시간표(반 기준 고정)
  const mandatoryRows = lec.filter(r => norm(r["유형"]) === "필수");
  const mandatoryByHome = new Map(); // home -> rows[]
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
      .sort((a, b) => (dayOrder.get(norm(a["요일"])) ?? 99) - (dayOrder.get(norm(b["요일"])) ?? 99) || periodNum(a["교시"]) - periodNum(b["교시"]))
      .map(r => `${norm(r["과목"])} ${norm(r["요일"])}${norm(r["교시"])}(${norm(r["강의실"])})`)
      .join(" / "));
  }

  // 2) 탐구 반 정보 구성 (소속관/지역 기반)
  const exploreRows = lec.filter(r => norm(r["유형"]) === "탐구");
  const exploreBySection = new Map(); // section -> rows[]
  for (const r of exploreRows) {
    const sec = norm(r["반"]);
    if (!exploreBySection.has(sec)) exploreBySection.set(sec, []);
    exploreBySection.get(sec).push({
      ...r,
      소속관: norm(r["소속관"]),
      지역: norm(r["지역"]),
    });
  }

  const sections = new Map(); // sec -> info
  for (const [sec, rows] of exploreBySection.entries()) {
    const buildings = [...new Set(rows.map(x => norm(x["소속관"])).filter(Boolean))];
    const regions = [...new Set(rows.map(x => norm(x["지역"])).filter(Boolean))];

    sections.set(sec, {
      sec,
      buildings, // 소속관 리스트
      regions,   // 지역 리스트
      slots: slotsFromRows(rows),
      detail: buildDetail(rows), // 실제 강의실 표시
      capacity: minCapacity(rows),
    });
  }

  // 정원 카운트
  const counts = new Map([...sections.keys()].map(s => [s, 0]));

  function remainingCap(sec) {
    const cap = sections.get(sec).capacity;
    if (cap === null || cap === undefined || Number.isNaN(Number(cap))) return Infinity;
    return Number(cap) - (counts.get(sec) ?? 0);
  }
  function canFit(sec, n) {
    return remainingCap(sec) >= n;
  }

  function getCandidates(subject, building, region, blockedSlots, groupSize) {
    const s = norm(subject);
    if (!s) return [];

    const out = [];
    for (const [sec, info] of sections.entries()) {
      if (!norm(sec).startsWith(s)) continue;

      // ✅ 지역 제약
      if (region && info.regions.length > 0 && !info.regions.includes(region)) continue;

      // ✅ 소속관 제약(학생관 == 강의 소속관)
      if (building && info.buildings.length > 0 && !info.buildings.includes(building)) continue;

      if (!canFit(sec, groupSize)) continue;
      if (hasConflict(info.slots, blockedSlots)) continue;

      out.push(sec);
    }

    // 덜 찬 반 우선
    out.sort((a, b) => (counts.get(a) - counts.get(b)) || a.localeCompare(b));
    return out;
  }

  function scorePair(sec1, sec2, groupSize) {
    const c1 = (counts.get(sec1) ?? 0) + groupSize;
    const c2 = (counts.get(sec2) ?? 0) + groupSize;
    const maxC = Math.max(c1, c2);
    const sumC = c1 + c2;
    return [maxC, sumC, sec1, sec2];
  }

  // 3) 조합 그룹: 학생반+관+탐구1+탐구2
  const groups = new Map();
  for (const s of stuAll) {
    const home = norm(s["반"]);
    const building = norm(s["관"]);
    const t1 = norm(s["탐구1"]);
    const t2 = norm(s["탐구2"]);
    const key = `${home}||${building}||${t1}||${t2}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(s);
  }

  // 큰 그룹부터 처리
  const groupEntries = [...groups.entries()]
    .map(([key, members]) => ({ key, members }))
    .sort((a, b) => b.members.length - a.members.length);

  const results = [];
  const groupPairMap = new Map(); // key -> {sec1, sec2}

  function assignGroup({ key, members }) {
    const [home, building, subj1, subj2] = key.split("||");
    const groupSize = members.length;
    const studentRegion = regionFromBuilding(building);

    const blockedBase = new Set();
    const mandSlots = mandatorySlotsByHome.get(home);
    if (mandSlots) for (const x of mandSlots) blockedBase.add(x);

    let warnAll = false;
    const notesBase = [];

    if (!mandSlots) {
      warnAll = true;
      notesBase.push(`필수시간표 없음(반=${home})`);
    }

    if (!studentRegion) {
      warnAll = true;
      notesBase.push(`관→지역 매핑 실패(관=${building})`);
    }

    const preferred = groupPairMap.get(key);

    const tryPreferred = (pair, size) => {
      if (!pair) return null;
      const { sec1, sec2 } = pair;
      if (!sec1 || !sec2) return null;
      if (!sections.has(sec1) || !sections.has(sec2)) return null;
      if (!canFit(sec1, size) || !canFit(sec2, size)) return null;

      const i1 = sections.get(sec1);
      const i2 = sections.get(sec2);

      // 지역/소속관 확인
      if (studentRegion && i1.regions.length && !i1.regions.includes(studentRegion)) return null;
      if (studentRegion && i2.regions.length && !i2.regions.includes(studentRegion)) return null;
      if (building && i1.buildings.length && !i1.buildings.includes(building)) return null;
      if (building && i2.buildings.length && !i2.buildings.includes(building)) return null;

      if (hasConflict(i1.slots, blockedBase)) return null;

      const blocked = new Set(blockedBase);
      for (const x of i1.slots) blocked.add(x);
      if (hasConflict(i2.slots, blocked)) return null;

      return { sec1, sec2 };
    };

    // 1) preferred로 그룹 전체 시도
    const okPreferred = tryPreferred(preferred, groupSize);
    if (okPreferred) {
      counts.set(okPreferred.sec1, (counts.get(okPreferred.sec1) ?? 0) + groupSize);
      counts.set(okPreferred.sec2, (counts.get(okPreferred.sec2) ?? 0) + groupSize);
      return [{ size: groupSize, sec1: okPreferred.sec1, sec2: okPreferred.sec2, warn: warnAll, notes: notesBase }];
    }

    // 2) 그룹 전체 수용 가능한 pair 탐색
    const cand1 = getCandidates(subj1, building, studentRegion, blockedBase, groupSize);
    let best = null;

    for (const sec1 of cand1) {
      const blocked = new Set(blockedBase);
      for (const x of sections.get(sec1).slots) blocked.add(x);

      const cand2 = getCandidates(subj2, building, studentRegion, blocked, groupSize);
      for (const sec2 of cand2) {
        const sc = scorePair(sec1, sec2, groupSize);
        if (!best || sc < best.score) best = { sec1, sec2, score: sc };
      }
    }

    if (best) {
      groupPairMap.set(key, { sec1: best.sec1, sec2: best.sec2 });
      counts.set(best.sec1, (counts.get(best.sec1) ?? 0) + groupSize);
      counts.set(best.sec2, (counts.get(best.sec2) ?? 0) + groupSize);
      return [{ size: groupSize, sec1: best.sec1, sec2: best.sec2, warn: warnAll, notes: notesBase }];
    }

    // 3) 그룹 전체 불가 → 분할(최소 분할): 가능한 최대 chunk부터
    let remaining = groupSize;
    const chunks = [];
    let split = false;

    while (remaining > 0) {
      let found = null;

      for (let chunkSize = remaining; chunkSize >= 1; chunkSize--) {
        const okPref = tryPreferred(preferred, chunkSize);
        if (okPref) {
          found = { size: chunkSize, sec1: okPref.sec1, sec2: okPref.sec2 };
          break;
        }

        const c1 = getCandidates(subj1, building, studentRegion, blockedBase, chunkSize);
        let bestLocal = null;

        for (const sec1 of c1) {
          const blocked = new Set(blockedBase);
          for (const x of sections.get(sec1).slots) blocked.add(x);

          const c2 = getCandidates(subj2, building, studentRegion, blocked, chunkSize);
          for (const sec2 of c2) {
            const sc = scorePair(sec1, sec2, chunkSize);
            if (!bestLocal || sc < bestLocal.score) bestLocal = { sec1, sec2, score: sc };
          }
        }

        if (bestLocal) {
          found = { size: chunkSize, sec1: bestLocal.sec1, sec2: bestLocal.sec2 };
          break;
        }
      }

      if (!found) {
        split = true;
        chunks.push({
          size: remaining,
          sec1: "",
          sec2: "",
          warn: true,
          notes: [...notesBase, "조합 단위 배정 실패(지역/소속관/정원/충돌/개설없음)"]
        });
        remaining = 0;
        break;
      }

      if (found.size !== groupSize) split = true;

      counts.set(found.sec1, (counts.get(found.sec1) ?? 0) + found.size);
      counts.set(found.sec2, (counts.get(found.sec2) ?? 0) + found.size);

      chunks.push({
        size: found.size,
        sec1: found.sec1,
        sec2: found.sec2,
        warn: warnAll || split,
        notes: split ? [...notesBase, "⚠️ 같은 조합 그룹이 정원/제약으로 분할 배정됨"] : notesBase
      });

      remaining -= found.size;
    }

    const allAssignedOneChunk =
      chunks.length === 1 &&
      chunks[0].sec1 && chunks[0].sec2 &&
      chunks[0].size === groupSize;

    if (allAssignedOneChunk) {
      groupPairMap.set(key, { sec1: chunks[0].sec1, sec2: chunks[0].sec2 });
    }

    return chunks;
  }

  // 4) 배정 실행
  for (const g of groupEntries) {
    const chunks = assignGroup(g);

    let idx = 0;
    for (const ch of chunks) {
      const part = g.members.slice(idx, idx + ch.size);
      idx += ch.size;

      for (const stu of part) {
        const id = norm(stu["일련번호"]);
        const home = norm(stu["반"]);
        const building = norm(stu["관"]);
        const subj1 = norm(stu["탐구1"]);
        const subj2 = norm(stu["탐구2"]);

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
          // 디버그용(원치 않으면 제거 가능)
          "조합그룹키(참고)": g.key,
        });
      }
    }
  }

  // 5) 반별명단
  const roster = [];
  for (const [sec, info] of sections.entries()) {
    const members = results.filter(r => r["탐구1_배정반"] === sec || r["탐구2_배정반"] === sec);
    if (members.length === 0) continue;

    for (const m of members) {
      roster.push({
        "탐구반": sec,
        "지역": info.regions.join("/"),
        "소속관": info.buildings.join("/"),
        "일련번호": m["일련번호"],
        "학생반": m["학생반"],
        "탐구1": m["탐구1"],
        "탐구2": m["탐구2"],
        "정원": info.capacity ?? "",
        "탐구반_시간표": info.detail,
      });
    }
  }

  // 6) 요약
  const summary = [];
  for (const [sec, info] of sections.entries()) {
    const c = counts.get(sec) ?? 0;
    if (c <= 0) continue;
    summary.push({
      "탐구반": sec,
      "지역": info.regions.join("/"),
      "소속관": info.buildings.join("/"),
      "배정인원": c,
      "정원": info.capacity ?? "",
      "시간표": info.detail,
    });
  }
  summary.sort((a, b) => b["배정인원"] - a["배정인원"] || norm(a["탐구반"]).localeCompare(norm(b["탐구반"])));

  // 7) Output workbook 생성
  const outWb = new ExcelJS.Workbook();
  outWb.created = new Date();

  function addSheet(name, rows, warnRed = false) {
    const ws = outWb.addWorksheet(name);
    if (!rows || rows.length === 0) {
      ws.addRow(["데이터 없음"]);
      return ws;
    }

    const cols = Object.keys(rows[0]);
    ws.addRow(cols);

    // header style
    const header = ws.getRow(1);
    header.font = { bold: true, color: { argb: "FFFFFFFF" } };
    header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E79" } };
    header.alignment = { vertical: "middle", horizontal: "center" };

    for (const r of rows) ws.addRow(cols.map(c => r[c]));

    ws.views = [{ state: "frozen", ySplit: 1 }];

    // widths
    cols.forEach((c, idx) => {
      ws.getColumn(idx + 1).width = Math.min(Math.max(12, c.length + 2), 60);
    });

    // warning highlight
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

  // (선택) 디버그 컬럼을 빼고 싶으면 여기서 results에서 해당 키 삭제하거나, addSheet 전에 cols 필터링 방식으로 처리 가능

  const buf = await outWb.xlsx.writeBuffer();
  return buf;
}

function downloadBuffer(buf, filename) {
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/** =========================
 * UI 이벤트
 ========================= */
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
