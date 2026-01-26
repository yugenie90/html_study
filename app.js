/* global ExcelJS */

const $file = document.getElementById("file");
const $run = document.getElementById("run");
const $download = document.getElementById("download");
const $log = document.getElementById("log");

let inputArrayBuffer = null;
let outputArrayBuffer = null;

function log(msg) {
  $log.textContent += msg + "\n";
}

function norm(s) {
  return (s ?? "").toString().trim();
}

function isExcluded(text) {
  const s = norm(text);
  return s.includes("목동") || s.includes("기숙");
}

function buildingOfRoom(room) {
  const s = norm(room);
  if (!s) return "";
  return s.split(/\s+/)[0]; // "W관"
}

function periodNum(p) {
  // "5교시" -> 5
  const m = norm(p).match(/(\d+)/);
  return m ? parseInt(m[1], 10) : 99;
}

const dayOrder = new Map([
  ["월요일", 1], ["화요일", 2], ["수요일", 3], ["목요일", 4], ["금요일", 5], ["토요일", 6], ["일요일", 7],
]);

async function readWorkbook(arrayBuffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(arrayBuffer);
  return wb;
}

function sheetToObjects(ws) {
  const headerRow = ws.getRow(1);
  const headers = [];
  headerRow.eachCell((cell, col) => headers[col] = norm(cell.value));

  const rows = [];
  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const obj = {};
    let empty = true;
    headers.forEach((h, col) => {
      if (!h) return;
      const v = row.getCell(col).value;
      const val = (v && typeof v === "object" && "text" in v) ? v.text : v;
      obj[h] = val;
      if (norm(val) !== "") empty = false;
    });
    if (!empty) rows.push(obj);
  });
  return rows;
}

function hasConflict(slotsA, slotsB) {
  // slots: Set("월요일|5교시")
  for (const s of slotsA) {
    if (slotsB.has(s)) return true;
  }
  return false;
}

function slotsFromRows(rows) {
  const set = new Set();
  for (const r of rows) {
    const key = `${norm(r["요일"])}|${norm(r["교시"])}`;
    set.add(key);
  }
  return set;
}

function buildDetail(rows) {
  // "수요일5교시(W관 701호) / 수요일6교시(W관 701호)"
  const sorted = [...rows].sort((a, b) => {
    const da = dayOrder.get(norm(a["요일"])) ?? 99;
    const db = dayOrder.get(norm(b["요일"])) ?? 99;
    if (da !== db) return da - db;
    return periodNum(a["교시"]) - periodNum(b["교시"]);
  });
  return sorted.map(r => `${norm(r["요일"])}${norm(r["교시"])}(${norm(r["강의실"])})`).join(" / ");
}

function minCapacity(rows) {
  // rows may have blank; take min numeric
  let cap = null;
  for (const r of rows) {
    const v = r["최대인원"];
    const n = Number(v);
    if (!Number.isFinite(n)) continue;
    cap = (cap === null) ? n : Math.min(cap, n);
  }
  return cap;
}

function startsWithSubject(sectionName, subjectCode) {
  return norm(sectionName).startsWith(norm(subjectCode));
}

async function runAssignment(inputWb) {
  const lecWs = inputWb.getWorksheet("강의정보");
  const stuWs = inputWb.getWorksheet("학생정보");
  if (!lecWs || !stuWs) {
    throw new Error("시트 이름이 정확해야 합니다: 강의정보, 학생정보");
  }

  const lecAll = sheetToObjects(lecWs);
  const stuAll = sheetToObjects(stuWs);

  log(`강의정보 ${lecAll.length}행, 학생정보 ${stuAll.length}행 로드`);

  // 1) 강의정보 필터: 목동/기숙 제외 + 유형=선택 제외
  const lec = lecAll.filter(r => {
    if (norm(r["유형"]) === "선택") return false;
    if (isExcluded(r["강의실"]) || isExcluded(r["과목"]) || isExcluded(r["요일&교시&강의실"])) return false;
    return true;
  });

  // 2) 필수 시간표(반 기준 고정)
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
      .sort((a,b)=> (dayOrder.get(norm(a["요일"]))??99)-(dayOrder.get(norm(b["요일"]))??99) || periodNum(a["교시"])-periodNum(b["교시"]))
      .map(r => `${norm(r["과목"])} ${norm(r["요일"])}${norm(r["교시"])}(${norm(r["강의실"])})`)
      .join(" / "));
  }

  // 3) 탐구 반 정보 구성
  const exploreRows = lec.filter(r => norm(r["유형"]) === "탐구");
  const exploreBySection = new Map(); // section -> rows[]
  for (const r of exploreRows) {
    const sec = norm(r["반"]);
    if (!exploreBySection.has(sec)) exploreBySection.set(sec, []);
    const rr = { ...r, 관: buildingOfRoom(r["강의실"]) };
    exploreBySection.get(sec).push(rr);
  }

  const sections = new Map(); // sec -> info
  for (const [sec, rows] of exploreBySection.entries()) {
    const buildings = [...new Set(rows.map(r => norm(r["관"])).filter(Boolean))];
    sections.set(sec, {
      sec,
      buildings,
      slots: slotsFromRows(rows),
      detail: buildDetail(rows),
      capacity: minCapacity(rows), // may be null
    });
  }

  // 카운트(정원)
  const counts = new Map([...sections.keys()].map(s => [s, 0]));

  function hasCap(sec) {
    const cap = sections.get(sec).capacity;
    if (cap === null || cap === undefined) return true;
    return (counts.get(sec) ?? 0) < cap;
  }

  // 같은 반이면 같은 탐구반 우선: (학생반, 관, 과목) -> sec
  const preferMap = new Map();

  function getCandidates(subject, building) {
    const s = norm(subject);
    if (!s) return [];
    const out = [];
    for (const [sec, info] of sections.entries()) {
      if (!startsWithSubject(sec, s)) continue;
      if (building && !info.buildings.includes(building)) continue;
      out.push(sec);
    }
    // 기본: 인원 적은 반 우선
    out.sort((a,b)=> (counts.get(a)-counts.get(b)) || a.localeCompare(b));
    return out;
  }

  function pickSection({home, building, subject, blockedSlots, preferKey}) {
    const preferred = preferMap.get(preferKey);
    if (preferred && sections.has(preferred)) {
      const info = sections.get(preferred);
      if (hasCap(preferred) && !hasConflict(info.slots, blockedSlots)) {
        return preferred;
      }
    }

    const candidates = getCandidates(subject, building)
      .filter(sec => hasCap(sec) && !hasConflict(sections.get(sec).slots, blockedSlots));

    if (candidates.length === 0) return "";
    return candidates[0];
  }

  const results = [];
  for (const stu of stuAll) {
    const id = norm(stu["일련번호"]);
    const building = norm(stu["관"]);   // W관
    const home = norm(stu["반"]);       // D(1)
    const subj1 = norm(stu["탐구1"]);
    const subj2 = norm(stu["탐구2"]);

    let warn = false;
    const notes = [];

    const blocked = new Set();
    const mandSlots = mandatorySlotsByHome.get(home);
    if (mandSlots) {
      for (const x of mandSlots) blocked.add(x);
    } else {
      warn = true;
      notes.push(`필수시간표 없음(반=${home})`);
    }

    // 탐구1
    const key1 = `${home}|${building}|${subj1}|T1`;
    const sec1 = pickSection({
      home, building, subject: subj1, blockedSlots: blocked, preferKey: key1
    });

    if (!sec1) {
      warn = true;
      notes.push(`탐구1(${subj1}) 배정불가(관/정원/필수충돌/개설없음)`);
    } else {
      preferMap.set(key1, sec1);
      counts.set(sec1, (counts.get(sec1) ?? 0) + 1);
      for (const x of sections.get(sec1).slots) blocked.add(x);
    }

    // 탐구2 (탐구1 + 필수 포함한 blocked 기준으로 충돌 금지)
    const key2 = `${home}|${building}|${subj2}|T2`;
    const sec2 = pickSection({
      home, building, subject: subj2, blockedSlots: blocked, preferKey: key2
    });

    if (!sec2) {
      warn = true;
      notes.push(`탐구2(${subj2}) 배정불가(관/정원/충돌/개설없음)`);
    } else {
      preferMap.set(key2, sec2);
      counts.set(sec2, (counts.get(sec2) ?? 0) + 1);
      for (const x of sections.get(sec2).slots) blocked.add(x);
    }

    results.push({
      일련번호: id,
      관: building,
      학생반: home,
      탐구1: subj1,
      탐구1_배정반: sec1,
      탐구1_시간표: sec1 ? sections.get(sec1).detail : "",
      탐구2: subj2,
      탐구2_배정반: sec2,
      탐구2_시간표: sec2 ? sections.get(sec2).detail : "",
      경고: warn ? "⚠️" : "",
      비고: notes.join(" | "),
      "필수시간표(참고)": mandatoryDetailByHome.get(home) ?? "",
    });
  }

  // 반별명단
  const roster = [];
  for (const [sec, info] of sections.entries()) {
    const members = results.filter(r => r.탐구1_배정반 === sec || r.탐구2_배정반 === sec);
    if (members.length === 0) continue;
    for (const m of members) {
      roster.push({
        탐구반: sec,
        관: info.buildings.join("/"),
        일련번호: m.일련번호,
        학생반: m.학생반,
        탐구1: m.탐구1,
        탐구2: m.탐구2,
        정원: info.capacity ?? "",
        탐구반_시간표: info.detail,
      });
    }
  }

  // 요약
  const summary = [];
  for (const [sec, info] of sections.entries()) {
    const c = counts.get(sec) ?? 0;
    if (c <= 0) continue;
    summary.push({
      탐구반: sec,
      관: info.buildings.join("/"),
      배정인원: c,
      정원: info.capacity ?? "",
      시간표: info.detail,
    });
  }
  summary.sort((a,b)=> b.배정인원 - a.배정인원 || a.탐구반.localeCompare(b.탐구반));

  // output workbook (스타일 포함)
  const outWb = new ExcelJS.Workbook();
  outWb.created = new Date();

  function addSheet(name, rows, warnRed = false) {
    const ws = outWb.addWorksheet(name);
    if (rows.length === 0) {
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

    // warning rows highlight
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

  const buf = await outWb.xlsx.writeBuffer();
  return buf;
}

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
  const name = `탐구반_자동배정_결과_${new Date().toISOString().slice(0,10)}.xlsx`;
  downloadBuffer(outputArrayBuffer, name);
});
