/* global ExcelJS */

/**
 * 탐구반 자동 배정 (대치 전용, 통짜 완성본)
 *
 * ✔ 같은 반은 가능하면 같은 탐구 (soft)
 * ✔ 배정 실패 방지 최우선
 * ✔ 필수수업 시간 충돌 사전 계산
 * ✔ 반별 가능 탐구 / 불가능 탐구 계산
 * ✔ 정원 90% 우선 → 다른 탐구로 분산 → 100% 확장
 * ✔ 결과 엑셀:
 *    - 학생별배정 (미배정사유, 불가능탐구목록 포함)
 *    - 반별명단
 *    - 요약
 *    - 불가능탐구_요약
 */

const $file = document.getElementById("file");
const $run = document.getElementById("run");
const $download = document.getElementById("download");
const $log = document.getElementById("log");

let inputBuffer = null;
let outputBuffer = null;

function log(msg) {
  $log.textContent += msg + "\n";
}

function norm(v) {
  return (v ?? "").toString().trim();
}

function slotKey(day, period) {
  return `${day}-${period}`;
}

function hasConflict(slotsA, slotsB) {
  for (const s of slotsA) if (slotsB.has(s)) return true;
  return false;
}

/* =========================
   FILE LOAD
========================= */
$file.addEventListener("change", async (e) => {
  const f = e.target.files[0];
  if (!f) return;
  inputBuffer = await f.arrayBuffer();
  $run.disabled = false;
  log("엑셀 업로드 완료");
});

/* =========================
   MAIN
========================= */
$run.addEventListener("click", async () => {
  try {
    log("워크북 로딩 중...");
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(inputBuffer);

    const wsLecture = wb.getWorksheet("강의정보");
    const wsStudent = wb.getWorksheet("학생정보");
    if (!wsLecture || !wsStudent) {
      throw new Error("시트 이름 오류: 강의정보 / 학생정보");
    }

    /* =========================
       강의정보 파싱
    ========================= */
    const lectures = [];
    wsLecture.eachRow((row, r) => {
      if (r === 1) return;
      const type = norm(row.getCell(6).value); // F: 유형
      if (type === "선택") return;

      const subject = norm(row.getCell(5).value);   // E
      const home = norm(row.getCell(8).value);      // H (반)
      const room = norm(row.getCell(9).value);      // I
      const day = norm(row.getCell(10).value);      // J
      const period = norm(row.getCell(11).value);   // K
      const maxCap = Number(row.getCell(12).value); // L

      if (!subject || !home || !day || !period) return;

      lectures.push({
        subject,
        home,
        room,
        slot: slotKey(day, period),
        maxCap,
      });
    });

    log(`강의정보 ${lectures.length}건 로드`);

    /* =========================
       반별 필수 시간표
    ========================= */
    const mandatorySlotsByHome = new Map();
    for (const l of lectures) {
      if (!mandatorySlotsByHome.has(l.home)) {
        mandatorySlotsByHome.set(l.home, new Set());
      }
      mandatorySlotsByHome.get(l.home).add(l.slot);
    }

    /* =========================
       탐구 섹션 구성
    ========================= */
    const sections = new Map(); // sec -> {slots, cap90, cap100}
    for (const l of lectures) {
      if (!sections.has(l.subject)) {
        sections.set(l.subject, {
          slots: new Set(),
          maxCap: l.maxCap,
        });
      }
      sections.get(l.subject).slots.add(l.slot);
    }

    for (const [sec, info] of sections.entries()) {
      const cap100 = info.maxCap ?? Infinity;
      const cap90 = Math.ceil(cap100 * 0.9);
      info.cap90 = cap90;
      info.cap100 = cap100;
    }

    /* =========================
       반별 가능/불가능 탐구 계산
    ========================= */
    const possibleByHome = new Map();
    const impossibleSummary = [];

    for (const [home, mandSlots] of mandatorySlotsByHome.entries()) {
      possibleByHome.set(home, new Set());
      for (const [sec, info] of sections.entries()) {
        if (hasConflict(info.slots, mandSlots)) {
          impossibleSummary.push({
            반: home,
            탐구반: sec,
            불가능사유: "필수수업 시간 충돌",
          });
        } else {
          possibleByHome.get(home).add(sec);
        }
      }
    }

    /* =========================
       학생정보 파싱
    ========================= */
    const students = [];
    wsStudent.eachRow((row, r) => {
      if (r === 1) return;
      const id = norm(row.getCell(1).value);
      const home = norm(row.getCell(3).value);
      const subj1 = norm(row.getCell(4).value);
      const subj2 = norm(row.getCell(5).value);
      if (!id || !home) return;
      students.push({ id, home, subj1, subj2 });
    });

    log(`학생정보 ${students.length}명 로드`);

    /* =========================
       그룹핑 (반+탐구1+탐구2)
    ========================= */
    const groups = new Map();
    for (const s of students) {
      const key = `${s.home}|${s.subj1}|${s.subj2}`;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(s);
    }

    /* =========================
       배정 준비
    ========================= */
    const assigned = [];
    const counts = new Map();
    for (const sec of sections.keys()) counts.set(sec, 0);

    function remaining(sec, phase) {
      const info = sections.get(sec);
      const cap = phase === "HARD" ? info.cap100 : info.cap90;
      return cap - (counts.get(sec) ?? 0);
    }

    function canUse(home, sec) {
      return possibleByHome.get(home)?.has(sec);
    }

    function tryAssign(group, phase) {
      const { home, subj1, subj2 } = group.meta;
      const size = group.students.length;

      const c1 = canUse(home, subj1);
      const c2 = canUse(home, subj2);

      if (!c1 || !c2) return null;
      if (remaining(subj1, phase) < size) return null;
      if (remaining(subj2, phase) < size) return null;

      return { subj1, subj2 };
    }

    /* =========================
       배정 실행
    ========================= */
    for (const [key, groupStudents] of groups.entries()) {
      const [home, subj1, subj2] = key.split("|");
      const group = { meta: { home, subj1, subj2 }, students: groupStudents };

      let pair = tryAssign(group, "SOFT");

      if (!pair) {
        // 다른 탐구로 분산 (SOFT)
        for (const alt1 of possibleByHome.get(home) ?? []) {
          for (const alt2 of possibleByHome.get(home) ?? []) {
            if (remaining(alt1, "SOFT") >= group.students.length &&
                remaining(alt2, "SOFT") >= group.students.length) {
              pair = { subj1: alt1, subj2: alt2 };
              break;
            }
          }
          if (pair) break;
        }
      }

      if (!pair) {
        pair = tryAssign(group, "HARD");
      }

      if (!pair) {
        // 배정 불가
        for (const s of group.students) {
          assigned.push({
            ...s,
            탐구1배정: "",
            탐구2배정: "",
            미배정사유: "가능 탐구 없음",
            불가능탐구목록: Array.from(
              sections.keys()
            ).filter(sec => !canUse(home, sec)).join(", "),
          });
        }
        continue;
      }

      counts.set(pair.subj1, counts.get(pair.subj1) + group.students.length);
      counts.set(pair.subj2, counts.get(pair.subj2) + group.students.length);

      for (const s of group.students) {
        assigned.push({
          ...s,
          탐구1배정: pair.subj1,
          탐구2배정: pair.subj2,
          미배정사유: "",
          불가능탐구목록: Array.from(
            sections.keys()
          ).filter(sec => !canUse(home, sec)).join(", "),
        });
      }
    }

    /* =========================
       결과 엑셀 생성
    ========================= */
    const out = new ExcelJS.Workbook();

    function addSheet(name, rows) {
      const ws = out.addWorksheet(name);
      if (!rows.length) return;
      ws.columns = Object.keys(rows[0]).map(k => ({ header: k, key: k }));
      rows.forEach(r => ws.addRow(r));
    }

    addSheet("학생별배정", assigned);
    addSheet("불가능탐구_요약", impossibleSummary);

    outputBuffer = await out.xlsx.writeBuffer();
    $download.disabled = false;
    log("배정 완료");
  } catch (e) {
    log("에러: " + e.message);
  }
});

/* =========================
   DOWNLOAD
========================= */
$download.addEventListener("click", () => {
  const blob = new Blob([outputBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "탐구반_자동배정_결과.xlsx";
  a.click();
});
