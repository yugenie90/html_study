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
 */
const SUBJECT_MAP = new Map([
  ["생윤", "생활과윤리"], ["사문", "사회문화"], ["윤사", "윤리와사상"],
  ["정법", "정치와법"], ["세지", "세계지리"], ["한지", "한국지리"],
  ["동사", "동아시아사"], ["세사", "세계사"], ["경제", "경제"],
  ["물1", "물리학1"], ["화1", "화학1"], ["생1", "생명과학1"], ["지1", "지구과학1"],
]);

/** 강의정보 풀네임 prefix → 학생 약어로 역변환(출력용) */
const PREFIX_TO_ABBR = [...SUBJECT_MAP.entries()].map(([abbr, full]) => [full, abbr])
  .sort((a, b) => b[0].length - a[0].length); // 긴 prefix 우선 매칭

function subjectPrefixFromStudent(s) {
  const key = norm(s);
  if (!key) return "";
  return SUBJECT_MAP.get(key) ?? key;
}

function toAbbrevSectionName(secName) {
  const s = norm(secName);
  if (!s) return "";
  for (const [full, abbr] of PREFIX_TO_ABBR) {
    if (s.startsWith(full)) return abbr + s.slice(full.length);
  }
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
    if (!wsLecture || !wsStudent) throw new Error("시트 이름 오류: 강의정보 / 학생정보");

    // =========================
    // 1) 강의정보 파싱
    // =========================
    const mandatorySlotsByHome = new Map();
    const sections = new Map();
    wsLecture.eachRow((row, r) => {
      if (r === 1) return;
      const type = norm(row.getCell(6).value);
      if (type !== "필수" && type !== "탐구") return;
      const subject = norm(row.getCell(5).value);
      const home = norm(row.getCell(8).value);
      const day = norm(row.getCell(10).value);
      const period = norm(row.getCell(11).value);
      if (!subject || !home || !day || !period) return;
      const slot = slotKey(day, period);
      if (type === "필수") {
        if (!mandatorySlotsByHome.has(home)) mandatorySlotsByHome.set(home, new Set());
        mandatorySlotsByHome.get(home).add(slot);
      } else {
        if (!sections.has(subject)) {
          const maxCapRaw = row.getCell(12).value;
          const cap100 = Number(maxCapRaw);
          const safeCap100 = Number.isFinite(cap100) ? cap100 : Infinity;
          sections.set(subject, {
            slots: new Set(),
            cap100: safeCap100,
            cap90: Math.ceil(safeCap100 * 0.9),
            name: subject,
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
      const id = norm(row.getCell(1).value);
      const home = norm(row.getCell(3).value);
      const subj1 = norm(row.getCell(4).value);
      const subj2 = norm(row.getCell(5).value);
      if (!id || !home) return;
      students.push({ id, home, subj1, subj2 });
    });
    log(`학생정보 ${students.length}명 로드`);

    // =========================
    // 3) 반별 가능/불가능 탐구 계산 + 요약
    // =========================
    const possibleByHome = new Map();
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
    // 4) 배정 준비 (헬퍼 함수)
    // =========================
    const counts = new Map();
    for (const sec of sections.keys()) counts.set(sec, 0);
    const failureReasons = new Map(); // secName -> { reason: count }

    function remaining(sec, capType) {
      const info = sections.get(sec);
      const cap = capType === "HARD" ? info.cap100 : info.cap90;
      return cap - (counts.get(sec) ?? 0);
    }
    function canUse(home, sec) { return possibleByHome.get(home)?.has(sec); }
    function allOpenedSectionsByPrefix(prefix) {
      if (!prefix) return [];
      return [...sections.keys()].filter(sec => sec.startsWith(prefix));
    }
    function candidatesByPrefix(home, prefix, capType, needCount, extraBlockedSlots) {
      const out = [];
      const blocked = extraBlockedSlots ?? new Set();
      for (const [sec, info] of sections.entries()) {
        if (!sec.startsWith(prefix)) continue;
        if (!canUse(home, sec)) continue;
        if (remaining(sec, capType) < needCount) continue;
        if (hasConflict(info.slots, blocked)) continue;
        out.push(sec);
      }
      return out.sort((a, b) => (counts.get(a) - counts.get(b)) || a.localeCompare(b));
    }

    // =========================
    // 5) 그룹핑 및 우선순위 정렬
    // =========================
    log("그룹별 배정 가능 조합 계산 시작...");
    const groupsMap = new Map();
    students.forEach(s => {
      const key = `${s.home}|${s.subj1}|${s.subj2}`;
      if (!groupsMap.has(key)) groupsMap.set(key, []);
      groupsMap.get(key).push(s);
    });

    const groups = [...groupsMap.entries()].map(([key, groupStudents]) => {
      const [home, want1, want2] = key.split("|");
      return {
        key, students: groupStudents, home, want1, want2,
        prefix1: subjectPrefixFromStudent(want1),
        prefix2: subjectPrefixFromStudent(want2),
        size: groupStudents.length,
        softOptions: 0, hardOptions: 0, reason: ""
      };
    });

    groups.forEach(g => {
      ensureHome(g.home);
      const mand = mandatorySlotsByHome.get(g.home);
      if (!allOpenedSectionsByPrefix(g.prefix1).length) g.reason = "탐구1 미개설";
      if (!allOpenedSectionsByPrefix(g.prefix2).length) g.reason += (g.reason ? " " : "") + "탐구2 미개설";
      if (g.reason) return;

      for (const capType of ["SOFT", "HARD"]) {
        const cand1 = candidatesByPrefix(g.home, g.prefix1, capType, g.size, mand);
        let optionsCount = 0;
        const validPairs = []; // 유일 페어를 찾기 위해 유효한 페어를 저장
        for (const sec1 of cand1) {
          const blocked2 = new Set([...mand, ...sections.get(sec1).slots]);
          const cand2 = candidatesByPrefix(g.home, g.prefix2, capType, g.size, blocked2);
          for (const sec2 of cand2) {
            validPairs.push({ sec1, sec2 });
          }
          optionsCount += cand2.length;
        }
        if (capType === "SOFT") {
          g.softOptions = optionsCount;
          g.softPairs = validPairs; // soft cap 기준 유효 페어 저장
        } else {
          g.hardOptions = optionsCount;
          g.hardPairs = validPairs; // hard cap 기준 유효 페어 저장
        }
      }
    });

    log(`그룹 총 ${groups.length}개 우선순위 계산 완료.`);

    let assigned = []; // <--- 선언 위치 이동
    // --- 선배정 (Forced First) 단계 ---
    log("선배정 (Forced First) 단계 시작...");
    const forcedFirstAssigned = [];
    const regularGroups = []; // 선배정 대상이 아닌 그룹들

    // 선배정 대상 식별 및 배정 (softOptions == 1)
    let softForcedCount = 0;
    let hardForcedCount = 0;

    for (const group of groups) {
      if (group.reason) { // 그룹 전체 미배정 (미개설 등)
        group.students.forEach(s => assigned.push({ ...s, 탐구1배정: "", 탐구2배정: "", 미배정사유: group.reason, 비고: "" }));
        continue;
      }
      
      let assignedByForcedFirst = false;
      let capTypeUsed = "";
      let chosenPair = null;

      if (group.softOptions === 1) {
        chosenPair = group.softPairs[0];
        capTypeUsed = "SOFT";
        softForcedCount++;
        assignedByForcedFirst = true;
      } else if (group.softOptions === 0 && group.hardOptions === 1) {
        chosenPair = group.hardPairs[0];
        capTypeUsed = "HARD";
        hardForcedCount++;
        assignedByForcedFirst = true;
      }

      if (assignedByForcedFirst && chosenPair) {
        const { sec1, sec2 } = chosenPair;
        const size = group.size;
        
        // 배정 반영
        counts.set(sec1, (counts.get(sec1) ?? 0) + size);
        counts.set(sec2, (counts.get(sec2) ?? 0) + size);
        
        group.students.forEach(s => assigned.push({
          ...s, 탐구1배정: toAbbrevSectionName(sec1), 탐구2배정: toAbbrevSectionName(sec2),
          미배정사유: "", 비고: capTypeUsed === "HARD" ? "선배정(정원 100% 허용)" : "선배정"
        }));
        forcedFirstAssigned.push(group);
      } else {
        regularGroups.push(group);
      }
    }
    log(`선배정 대상 개수: soft==1 (${softForcedCount}개), soft==0 && hard==1 (${hardForcedCount}개)`);
    if (forcedFirstAssigned.length > 0) {
      log("선배정된 그룹 상위 5개:");
      forcedFirstAssigned.slice(0, 5).forEach(g => {
        const assignedStudent = assigned.find(s => s.id === g.students[0].id);
        if (assignedStudent) {
            log(`- 반: ${g.home}, 탐구1: ${g.want1}, 탐구2: ${g.want2}, 인원: ${g.size}, 배정페어: ${assignedStudent.탐구1배정}/${assignedStudent.탐구2배정}`);
        }
      });
    }

    // 선배정된 그룹을 제외한 나머지 그룹으로 groups 배열 재구성
    groups.splice(0, groups.length, ...regularGroups); 

    groups.sort((a, b) =>
      // 0인 경우는 뒤로 (양수인 경우만 오름차순)
      (a.softOptions === 0 ? Infinity : a.softOptions) - (b.softOptions === 0 ? Infinity : b.softOptions) ||
      (a.hardOptions === 0 ? Infinity : a.hardOptions) - (b.hardOptions === 0 ? Infinity : b.hardOptions) ||
      b.size - a.size ||
      a.home.localeCompare(b.home)
    );
    // (디버그 로그 생략)

    // =========================
    // 6) ★★★ 배정 로직 ★★★
    // =========================
    log("그룹 배정 시작 (정렬된 순서 기반)...");
    const failureStats = new Map(); // For logging

    function tryAssignGroupAsPair(group, capType) {
      const { home, prefix1, prefix2, size, students: groupStudents } = group;
      const mand = mandatorySlotsByHome.get(home);
      const cand1 = candidatesByPrefix(home, prefix1, capType, size, mand);
      for (const sec1 of cand1) {
        const blocked2 = new Set([...mand, ...sections.get(sec1).slots]);
        const cand2 = candidatesByPrefix(home, prefix2, capType, size, blocked2);
        if (cand2.length > 0) {
          const sec2 = cand2[0];
          counts.set(sec1, (counts.get(sec1) ?? 0) + size);
          counts.set(sec2, (counts.get(sec2) ?? 0) + size);
          groupStudents.forEach(s => assigned.push({
            ...s, 탐구1배정: toAbbrevSectionName(sec1), 탐구2배정: toAbbrevSectionName(sec2),
            미배정사유: "", 비고: capType === "HARD" ? "정원 100% 허용" : ""
          }));
          return true;
        }
      }
      return false;
    }

    function getSingleAssignReason(home, prefix, blocked) {
        if (allOpenedSectionsByPrefix(prefix).length === 0) return "수업 미개설";
        const isPossible = allOpenedSectionsByPrefix(prefix).some(sec => canUse(home, sec));
        if (!isPossible) return "필수 시간 충돌";
        if (blocked && allOpenedSectionsByPrefix(prefix).some(sec => hasConflict(sections.get(sec).slots, blocked))) return "다른 탐구와 시간 충돌";
        return "정원 부족";
    }

    function recordFailure(subjectPrefix, reason) {
        if (!failureReasons.has(subjectPrefix)) {
            failureReasons.set(subjectPrefix, new Map());
        }
        const reasonMap = failureReasons.get(subjectPrefix);
        reasonMap.set(reason, (reasonMap.get(reason) ?? 0) + 1);
    }

    function tryPartialAssignment(student) {
        const { home, subj1, subj2 } = student;
        const prefix1 = subjectPrefixFromStudent(subj1);
        const prefix2 = subjectPrefixFromStudent(subj2);
        const mand = mandatorySlotsByHome.get(home);

        const res = { 탐구1배정: "", 탐구2배정: "", 비고: [] };
        let assignedSlots = new Set();
        
        const subjectsToTry = [
            { key: 1, subj: subj1, prefix: prefix1, cands: candidatesByPrefix(home, prefix1, "SOFT", 1, mand).length },
            { key: 2, subj: subj2, prefix: prefix2, cands: candidatesByPrefix(home, prefix2, "SOFT", 1, mand).length },
        ].sort((a,b) => a.cands - b.cands); // 희소 과목 우선

        let s1_res = null, s2_res = null;

        for(const current of subjectsToTry){
            let singleRes = null;
            for (const capType of ["SOFT", "HARD"]) {
                const cands = candidatesByPrefix(home, current.prefix, capType, 1, new Set([...mand, ...assignedSlots]));
                if (cands.length > 0) {
                    singleRes = { sec: cands[0], capType };
                    break;
                }
            }
            if(singleRes){
                counts.set(singleRes.sec, (counts.get(singleRes.sec) ?? 0) + 1);
                assignedSlots = new Set([...assignedSlots, ...sections.get(singleRes.sec).slots]);
                if(current.key === 1) s1_res = singleRes; else s2_res = singleRes;
            }
        }
        
        // 결과 정리
        if(s1_res) {
            res.탐구1배정 = toAbbrevSectionName(s1_res.sec);
            if(s1_res.capType === 'HARD') res.비고.push('정원 100% 허용(탐구1)');
        }
        if(s2_res) {
            res.탐구2배정 = toAbbrevSectionName(s2_res.sec);
            if(s2_res.capType === 'HARD') res.비고.push('정원 100% 허용(탐구2)');
        }

        if(!s1_res && subj1) {
            const reason = getSingleAssignReason(home, prefix1, s2_res ? sections.get(s2_res.sec).slots : null);
            res.비고.push(`부분배정: 탐구1 미배정(${reason})`);
            recordFailure(prefix1, reason);
        }
        if(!s2_res && subj2) {
            const reason = getSingleAssignReason(home, prefix2, s1_res ? sections.get(s1_res.sec).slots : null);
            res.비고.push(`부분배정: 탐구2 미배정(${reason})`);
            recordFailure(prefix2, reason);
        }
        
        return res;
    }


    for (const group of groups) {
      if (group.reason) { // 그룹 전체 미배정 (미개설 등)
        group.students.forEach(s => {
          assigned.push({ ...s, 탐구1배정: "", 탐구2배정: "", 미배정사유: group.reason, 비고: "" });
          if(group.prefix1) recordFailure(group.prefix1, "수업 미개설");
          if(group.prefix2) recordFailure(group.prefix2, "수업 미개설");
        });
        continue;
      }
      if (tryAssignGroupAsPair(group, "SOFT")) continue;
      if (tryAssignGroupAsPair(group, "HARD")) continue;
      
      // 페어 실패 → 부분 배정 시도
      group.students.forEach(s => {
          const partialRes = tryPartialAssignment(s);
          const finalNote = partialRes.비고.join('; ');

          if (!partialRes.탐구1배정 && !partialRes.탐구2배정) {
              const reason1 = getSingleAssignReason(s.home, subjectPrefixFromStudent(s.subj1), null);
              const reason2 = getSingleAssignReason(s.home, subjectPrefixFromStudent(s.subj2), null);
              assigned.push({ ...s, 탐구1배정: "", 탐구2배정: "", 미배정사유: "페어/부분 배정 불가", 비고: finalNote });
              recordFailure(subjectPrefixFromStudent(s.subj1), `페어실패(${reason1}/${reason2})`);
              recordFailure(subjectPrefixFromStudent(s.subj2), `페어실패(${reason1}/${reason2})`);
          } else {
              assigned.push({ ...s, ...partialRes, 비고: finalNote, 미배정사유: "" });
          }
      });
    }

    // =========================
    // 7) 최종 검증 및 로그
    // =========================
    log("--- 최종 배정 통계 ---");
    const stats = { both: 0, s1_only: 0, s2_only: 0, none: 0 };
    assigned.forEach(r => {
      if (r.탐구1배정 && r.탐구2배정) stats.both++;
      else if (r.탐구1배정) stats.s1_only++;
      else if (r.탐구2배정) stats.s2_only++;
      else stats.none++;
    });
    log(`- 둘 다 배정: ${stats.both}명`);
    log(`- 탐구1만 배정: ${stats.s1_only}명`);
    log(`- 탐구2만 배정: ${stats.s2_only}명`);
    log(`- 둘 다 미배정: ${stats.none}명`);

    const remainingSections = [];
    log("\n--- 잔여 정원 분반 ---");
    [...sections.keys()].sort().forEach(sec => {
        const info = sections.get(sec);
        const currentCount = counts.get(sec) ?? 0;
        const hardCap = info.cap100;
        if (currentCount < hardCap) {
            const remainingSeats = hardCap - currentCount;
            log(`- ${sec}: ${remainingSeats}석 남음 (현재 ${currentCount}/${hardCap})`);
            remainingSections.push(sec);
        }
    });

    log("\n--- '정원 남는데 미배정' 분석 ---");
    const prefixesWithRemainingSeats = new Set(remainingSections.map(sec => {
        for (const [full, abbr] of PREFIX_TO_ABBR) {
            if (sec.startsWith(full)) return full;
        }
        return sec;
    }));

    for(const prefix of [...prefixesWithRemainingSeats].sort()) {
        if (failureReasons.has(prefix)) {
            const reasons = failureReasons.get(prefix);
            const top3 = [...reasons.entries()].sort((a,b) => b[1] - a[1]).slice(0, 3);
            if (top3.length > 0) {
                log(`\n# 과목: ${prefix} (잔여 정원 있음)`);
                log("  배정 실패 사유 TOP 3:");
                top3.forEach(([reason, count]) => {
                    log(`  - ${reason}: ${count}회`);
                });
            }
        }
    }

    // =========================
    // 8) 결과 엑셀 생성
    // =========================
    const out = new ExcelJS.Workbook();
    function addSheet(name, rows) {
      const ws = out.addWorksheet(name);
      if (!rows.length) return;
      ws.columns = Object.keys(rows[0]).map(k => ({ header: k, key: k }));
      ws.addRows(rows);
      ws.views = [{ state: "frozen", ySplit: 1 }];
    }
    addSheet("학생별배정", assigned);
    addSheet("불가능탐구_요약", impossibleSummary);
    outputBuffer = await out.xlsx.writeBuffer();
    $download.disabled = false;
    log("배정 완료. 다운로드 가능.");

  } catch (e) {
    console.error(e);
    log("에러: " + e.message);
  }
});

$download.addEventListener("click", () => {
  if (!outputBuffer) return;
  const blob = new Blob([outputBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "탐구반_자동배정_결과.xlsx";
  a.click();
});

document.querySelectorAll(".logic-toggle").forEach(btn => {
  btn.addEventListener("click", () => btn.closest(".logic-accordion").classList.toggle("open"));
});
