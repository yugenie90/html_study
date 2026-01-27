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

/**
 * 학생 약어 → 강의정보 과목 prefix(풀네임)로 변환
 * (강의정보 과목명이 "사회문화A" 같이 풀네임+반으로 들어오는 걸 전제)
 */
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

/** 강의정보 풀네임 prefix → 학생 약어로 역변환(출력용) */
const PREFIX_TO_ABBR = [
  ["생활과윤리", "생윤"],
  ["사회문화", "사문"],
  ["윤리와사상", "윤사"],
  ["정치와법", "정법"],
  ["세계지리", "세지"],
  ["한국지리", "한지"],
  ["동아시아사", "동사"],
  ["세계사", "세사"],
  ["경제", "경제"],
  ["물리학1", "물1"],
  ["화학1", "화1"],
  ["생명과학1", "생1"],
  ["지구과학1", "지1"],
].sort((a,b)=> b[0].length - a[0].length); // 긴 prefix 우선 매칭

function subjectPrefixFromStudent(s) {
  const key = norm(s);
  if (!key) return "";
  return SUBJECT_MAP.get(key) ?? key;
}

function toAbbrevSectionName(secName) {
  const s = norm(secName);
  if (!s) return "";
  for (const [full, abbr] of PREFIX_TO_ABBR) {
    if (s.startsWith(full)) {
      return abbr + s.slice(full.length); // A/B/C 같은 suffix 유지
    }
  }
  // 매칭 안되면 원본 반환(예외 케이스 대비)
  return s;
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

      const subject = norm(row.getCell(5).value); // E: 과목(탐구 섹션명: 사회문화A 등)
      const home = norm(row.getCell(8).value);    // H: 반
      const day = norm(row.getCell(10).value);    // J
      const period = norm(row.getCell(11).value); // K
      const maxCapRaw = row.getCell(12).value;    // L

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

    // =========================
    // 2) 학생정보 로드
    // =========================
    const students = [];
    wsStudent.eachRow((row, r) => {
      if (r === 1) return;
      const id = norm(row.getCell(1).value);     // A: 일련번호
      const home = norm(row.getCell(3).value);   // C: 반
      const subj1 = norm(row.getCell(4).value);  // D: 탐구1(약어)
      const subj2 = norm(row.getCell(5).value);  // E: 탐구2(약어)
      if (!id || !home) return;
      students.push({ id, home, subj1, subj2 });
    });

    log(`학생정보 ${students.length}명 로드`);

    // =========================
    // 3) 반별 가능/불가능 탐구 계산 + 불가능탐구_요약
    // =========================
    const possibleByHome = new Map(); // home -> Set(secName)
    const impossibleSummary = [];

    function ensureHome(home) {
      if (!mandatorySlotsByHome.has(home)) mandatorySlotsByHome.set(home, new Set());
      if (!possibleByHome.has(home)) possibleByHome.set(home, new Set());
    }

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
    // 4) 배정 준비
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

    // "해당 과목(prefix)"로 시작하는 모든 섹션(개설여부 확인용, home/충돌 무시)
    function allOpenedSectionsByPrefix(prefix) {
      const out = [];
      for (const sec of sections.keys()) {
        if (sec.startsWith(prefix)) out.push(sec);
      }
      return out;
    }

    // home에서 사용 가능 + prefix 일치 + 정원여유 + (sec1과 시간 충돌 금지 옵션) 후보
    function candidatesByPrefix(home, prefix, phase, needCount, extraBlockedSlots /* Set or null */) {
      const out = [];
      const blocked = extraBlockedSlots ?? new Set();

      for (const [sec, info] of sections.entries()) {
        if (!sec.startsWith(prefix)) continue;
        if (!canUse(home, sec)) continue;
        if (remaining(sec, phase) < needCount) continue;
        if (hasConflict(info.slots, blocked)) continue;
        out.push(sec);
      }

      // 덜 찬 섹션 우선
      out.sort((a,b)=> (counts.get(a)-counts.get(b)) || a.localeCompare(b));
      return out;
    }

    // =========================
    // 5) 그룹핑 (반 + 탐구1 + 탐구2): "가능하면" 같은 조합 유지
    // =========================
    const groups = new Map();
    for (const s of students) {
      const key = `${s.home}|${s.subj1}|${s.subj2}`;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(s);
    }

    const assigned = [];

    // 학생 1명 배정 (90% → 100%) / 반드시 같은 과목 prefix 안에서만
    function assignOneStudent(s) {
      const home = s.home;
      const prefix1 = subjectPrefixFromStudent(s.subj1);
      const prefix2 = subjectPrefixFromStudent(s.subj2);

      // 개설 여부 체크(아예 섹션이 없으면 미배정 표시)
      const opened1 = allOpenedSectionsByPrefix(prefix1);
      const opened2 = allOpenedSectionsByPrefix(prefix2);

      if (!prefix1 || opened1.length === 0) {
        return { ok: false, reason: "수업 미개설(탐구1)" };
      }
      if (!prefix2 || opened2.length === 0) {
        return { ok: false, reason: "수업 미개설(탐구2)" };
      }

      // 필수 슬롯 + 탐구1 슬롯과 탐구2 충돌 금지
      const mand = mandatorySlotsByHome.get(home) ?? new Set();

      // SOFT 먼저
      for (const phase of ["SOFT", "HARD"]) {
        const cand1 = candidatesByPrefix(home, prefix1, phase, 1, mand);
        for (const sec1 of cand1) {
          const blocked2 = new Set(mand);
          for (const sl of sections.get(sec1).slots) blocked2.add(sl);

          const cand2 = candidatesByPrefix(home, prefix2, phase, 1, blocked2);
          if (cand2.length === 0) continue;

          const sec2 = cand2[0];

          // 카운트 반영
          counts.set(sec1, (counts.get(sec1) ?? 0) + 1);
          counts.set(sec2, (counts.get(sec2) ?? 0) + 1);

          return { ok: true, sec1, sec2, phase };
        }
      }

      // 여기까지 왔다면: 같은 과목이지만 (필수충돌/정원/탐구1-2충돌)로 배정 불가
      return { ok: false, reason: "가능 탐구 없음(필수 충돌/정원 초과/탐구 간 충돌)" };
    }

    // =========================
    // 6) 그룹 단위로 "한 번에" 넣어보고, 안 되면 학생 단위로 분할
    // =========================
    for (const [key, groupStudents] of groups.entries()) {
      const [home, want1, want2] = key.split("|");
      ensureHome(home);

      const prefix1 = subjectPrefixFromStudent(want1);
      const prefix2 = subjectPrefixFromStudent(want2);

      // 개설 자체가 없으면: 그룹 전체 미배정(수업 미개설)
      const opened1 = allOpenedSectionsByPrefix(prefix1);
      const opened2 = allOpenedSectionsByPrefix(prefix2);

      if (!prefix1 || opened1.length === 0 || !prefix2 || opened2.length === 0) {
        const reason =
          (!prefix1 || opened1.length === 0) && (!prefix2 || opened2.length === 0)
            ? "수업 미개설(탐구1/탐구2)"
            : (!prefix1 || opened1.length === 0)
              ? "수업 미개설(탐구1)"
              : "수업 미개설(탐구2)";

        for (const s of groupStudents) {
          assigned.push({
            id: s.id,
            home: s.home,
            subj1: s.subj1,
            subj2: s.subj2,
            탐구1배정: "",
            탐구2배정: "",
            미배정사유: reason,
            비고: "",
          });
        }
        continue;
      }

      // 1) 그룹을 "같은 조합"으로 한 번에 배정 시도 (SOFT 먼저, 실패하면 HARD)
      let groupAssigned = false;
      let groupNote = "";

      const mand = mandatorySlotsByHome.get(home) ?? new Set();
      const size = groupStudents.length;

      function tryAssignGroup(phase) {
        // 탐구1 후보(그룹 size만큼 여유)
        const cand1 = candidatesByPrefix(home, prefix1, phase, size, mand);

        for (const sec1 of cand1) {
          const blocked2 = new Set(mand);
          for (const sl of sections.get(sec1).slots) blocked2.add(sl);

          const cand2 = candidatesByPrefix(home, prefix2, phase, size, blocked2);
          if (cand2.length === 0) continue;

          const sec2 = cand2[0];

          // 카운트 반영
          counts.set(sec1, (counts.get(sec1) ?? 0) + size);
          counts.set(sec2, (counts.get(sec2) ?? 0) + size);

          // 그룹 전체 배정
          for (const s of groupStudents) {
            assigned.push({
              id: s.id,
              home: s.home,
              subj1: s.subj1,
              subj2: s.subj2,
              탐구1배정: toAbbrevSectionName(sec1),
              탐구2배정: toAbbrevSectionName(sec2),
              미배정사유: "",
              비고: phase === "HARD" ? "정원 100% 허용" : "",
            });
          }
          return true;
        }
        return false;
      }

      // SOFT 그룹 배정 시도
      if (tryAssignGroup("SOFT")) {
        groupAssigned = true;
      } else if (tryAssignGroup("HARD")) {
        groupAssigned = true;
      }

      if (groupAssigned) continue;

      // 2) 그룹 단위 실패 → 같은 과목 안에서 학생 단위 분할 배정
      //    (여기서 절대 다른 과목으로는 안 넘어감)
      groupNote = "같은 반 조합 분할(정원/충돌)";

      for (const s of groupStudents) {
        const res = assignOneStudent(s);

        if (!res.ok) {
          assigned.push({
            id: s.id,
            home: s.home,
            subj1: s.subj1,
            subj2: s.subj2,
            탐구1배정: "",
            탐구2배정: "",
            미배정사유: res.reason,
            비고: "",
          });
          continue;
        }

        assigned.push({
          id: s.id,
          home: s.home,
          subj1: s.subj1,
          subj2: s.subj2,
          탐구1배정: toAbbrevSectionName(res.sec1),
          탐구2배정: toAbbrevSectionName(res.sec2),
          미배정사유: "",
          비고: res.phase === "HARD" ? `${groupNote} / 정원 100% 허용` : groupNote,
        });
      }
    }

    // =========================
    // 7) 결과 엑셀 생성
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
