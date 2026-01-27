/* global ExcelJS */

const $file = document.getElementById("file");
const $run = document.getElementById("run");
const $download = document.getElementById("download");
const $log = document.getElementById("log");

let inputBuffer = null;
let outputBuffer = null;

function log(msg) { $log.textContent += msg + "\n"; }
function norm(v) { return (v ?? "").toString().trim(); }
function slotKey(day, period) { return `${day}-${period}`; }

function hasConflict(slotsSet, blockedSet) {
  for (const s of slotsSet) if (blockedSet.has(s)) return true;
  return false;
}

// 학생 약칭 → 강의정보 과목 prefix 매칭
const SUBJECT_MAP = new Map([
  ["생윤", "생활과윤리"],
  ["사문", "사회문화"],
  ["윤사", "윤리와사상"],
  ["정법", "정치와법"],
  ["세지", "세계지리"],
  ["한지", "한국지리"],
  ["동사", "동아시아사"],
  ["세사", "세계사"],
  ["경제", "경제"],
  ["물1", "물리학1"],
  ["화1", "화학1"],
  ["생1", "생명과학1"],
  ["지1", "지구과학1"],
]);

function subjectPrefixFromStudent(s) {
  const key = norm(s);
  if (!key) return "";
  return SUBJECT_MAP.get(key) ?? key;
}

$file.addEventListener("change", async (e) => {
  const f = e.target.files[0];
  if (!f) return;
  inputBuffer = await f.arrayBuffer();
  $run.disabled = false;
  log("엑셀 업로드 완료");
});

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

    // =========================
    // 1) 강의정보 파싱: 필수 vs 탐구 분리
    // =========================
    const mandatorySlotsByHome = new Map(); // home -> Set(slot)
    const sections = new Map(); // secName -> {slots:Set, cap90, cap100}

    wsLecture.eachRow((row, r) => {
      if (r === 1) return;

      const type = norm(row.getCell(6).value);   // F: 유형 (필수/탐구/선택)
      if (type === "선택") return;

      const subject = norm(row.getCell(5).value); // E: 과목(탐구 섹션명)
      const home = norm(row.getCell(8).value);    // H: 반
      const day = norm(row.getCell(10).value);    // J: 요일
      const period = norm(row.getCell(11).value); // K: 교시
      const maxCapRaw = row.getCell(12).value;    // L: 최대인원

      if (!subject || !home || !day || !period) return;
      const slot = slotKey(day, period);

      if (type === "필수") {
        if (!mandatorySlotsByHome.has(home)) mandatorySlotsByHome.set(home, new Set());
        mandatorySlotsByHome.get(home).add(slot);
        return;
      }

      if (type === "탐구") {
        if (!sections.has(subject)) {
          const cap100 = Number(maxCapRaw);
          const safeCap100 = Number.isFinite(cap100) ? cap100 : Infinity;
          sections.set(subject, {
            slots: new Set(),
            cap100: safeCap100,
            cap90: Math.ceil(safeCap100 * 0.9),
          });
        }
        sections.get(subject).slots.add(slot);
      }
    });

    log(`탐구 섹션 ${sections.size}개 로드`);

    const possibleByHome = new Map(); // home -> Set(sec)
    const impossibleSummary = [];     // sheet rows

    function ensureHome(home) {
      if (!mandatorySlotsByHome.has(home)) mandatorySlotsByHome.set(home, new Set());
      if (!possibleByHome.has(home)) possibleByHome.set(home, new Set());
    }

    // =========================
    // 2) 학생정보 로드
    // =========================
    const students = [];
    wsStudent.eachRow((row, r) => {
      if (r === 1) return;
      const id = norm(row.getCell(1).value);     // A
      const home = norm(row.getCell(3).value);   // C
      const subj1 = norm(row.getCell(4).value);  // D
      const subj2 = norm(row.getCell(5).value);  // E
      if (!id || !home) return;
      ensureHome(home);
      students.push({ id, home, subj1, subj2 });
    });

    log(`학생정보 ${students.length}명 로드`);

    // =========================
    // 3) 반별 가능/불가능 탐구 계산
    // =========================
    for (const home of new Set(students.map(s => s.home))) {
      ensureHome(home);
      const mand = mandatorySlotsByHome.get(home);

      for (const [sec, info] of sections.entries()) {
        if (hasConflict(info.slots, mand)) {
          impossibleSummary.push({ 반: home, 탐구반: sec, 불가능사유: "필수수업 시간 충돌" });
        } else {
          possibleByHome.get(home).add(sec);
        }
      }
    }

    // =========================
    // 4) 그룹핑 (반 + 탐구1 + 탐구2)
    // =========================
    const groups = new Map();
    for (const s of students) {
      const key = `${s.home}|${s.subj1}|${s.subj2}`;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(s);
    }

    // =========================
    // 5) 배정 준비 (정원 90% 우선 → 분산 → 100%)
    // =========================
    const counts = new Map();
    for (const sec of sections.keys()) counts.set(sec, 0);

    function remaining(sec, phase) {
      const info = sections.get(sec);
      const cap = (phase === "HARD") ? info.cap100 : info.cap90;
      return cap - (counts.get(sec) ?? 0);
    }

    function canUse(home, sec) {
      return possibleByHome.get(home)?.has(sec);
    }

    function candidatesByPrefix(home, studentSubj, phase, needCount) {
      const prefix = subjectPrefixFromStudent(studentSubj);
      if (!prefix) return [];

      const out = [];
      for (const [sec] of sections.entries()) {
        if (!canUse(home, sec)) continue;
        if (!sec.startsWith(prefix)) continue;
        if (remaining(sec, phase) < needCount) continue;
        out.push(sec);
      }
      out.sort((a,b)=> (counts.get(a)-counts.get(b)) || a.localeCompare(b));
      return out;
    }

    const assigned = [];

    for (const [key, groupStudents] of groups.entries()) {
      const [home, want1, want2] = key.split("|");
      const size = groupStudents.length;

      let chosen = null; // {sec1, sec2}
      let note = "";     // 배정 이슈(비고)
      let phaseUsed = ""; // "SOFT" | "HARD" | ""

      // (1) 90%: 원하는 과목(prefix)로 시도
      const c1_soft = candidatesByPrefix(home, want1, "SOFT", size);
      const c2_soft = candidatesByPrefix(home, want2, "SOFT", size);

      outer1:
      for (const s1 of c1_soft) {
        for (const s2 of c2_soft) {
          chosen = { sec1: s1, sec2: s2 };
          phaseUsed = "SOFT";
          break outer1;
        }
      }

      // (2) 90%에서 막히면 분산(탐구 풀 내 아무 섹션)
      if (!chosen) {
        const pool = Array.from(possibleByHome.get(home) ?? [])
          .filter(sec => remaining(sec, "SOFT") > 0)
          .sort((a,b)=> (counts.get(a)-counts.get(b)) || a.localeCompare(b));

        outer2:
        for (const a of pool) {
          if (remaining(a, "SOFT") < size) continue;
          for (const b of pool) {
            if (remaining(b, "SOFT") < size) continue;
            chosen = { sec1: a, sec2: b };
            phaseUsed = "SOFT";
            note = "정원 90% 유지 위해 분산 배정";
            break outer2;
          }
        }
      }

      // (3) 그래도 없으면 100% 확장 (마지막 수단)
      if (!chosen) {
        const c1_hard = candidatesByPrefix(home, want1, "HARD", size);
        const c2_hard = candidatesByPrefix(home, want2, "HARD", size);

        outer3:
        for (const s1 of c1_hard) {
          for (const s2 of c2_hard) {
            chosen = { sec1: s1, sec2: s2 };
            phaseUsed = "HARD";
            note = "정원 100% 허용";
            break outer3;
          }
        }
      }

      // (4) 최종 실패: 미배정사유만 기록
      if (!chosen) {
        for (const s of groupStudents) {
          assigned.push({
            id: s.id,
            home: s.home,
            subj1: s.subj1,
            subj2: s.subj2,
            탐구1배정: "",
            탐구2배정: "",
            미배정사유: "가능 탐구 없음(필수 충돌/정원 초과)",
            비고: "",
          });
        }
        continue;
      }

      // 배정 반영
      counts.set(chosen.sec1, (counts.get(chosen.sec1) ?? 0) + size);
      counts.set(chosen.sec2, (counts.get(chosen.sec2) ?? 0) + size);

      // 배정 성공: 미배정사유는 항상 빈칸, 이슈는 비고에만
      for (const s of groupStudents) {
        assigned.push({
          id: s.id,
          home: s.home,
          subj1: s.subj1,
          subj2: s.subj2,
          탐구1배정: chosen.sec1,
          탐구2배정: chosen.sec2,
          미배정사유: "",
          비고: note,
        });
      }
    }

    // =========================
    // 6) 결과 엑셀 생성
    // =========================
    const out = new ExcelJS.Workbook();

    function addSheet(name, rows) {
      const ws = out.addWorksheet(name);
      if (!rows.length) return;
      ws.columns = Object.keys(rows[0]).map(k => ({ header: k, key: k }));
      rows.forEach(r => ws.addRow(r));
      ws.views = [{ state: "frozen", ySplit: 1 }];
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

$download.addEventListener("click", () => {
  const blob = new Blob([outputBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "탐구반_자동배정_결과.xlsx";
  a.click();
});
