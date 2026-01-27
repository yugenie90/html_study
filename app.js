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
      capacity: minCapacity(rows), // null 가능
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

  function getCandidates(subject, building, blockedSlots, groupSize) {
    const s = norm(subject);
    if (!s) return [];
    const out = [];
    for (const [sec, info] of sections.entries()) {
      if (!norm(sec).startsWith(s)) continue;
      if (building && !info.buildings.includes(building)) continue;
      if (!canFit(sec, groupSize)) continue;
      if (hasConflict(info.slots, blockedSlots)) continue;
      out.push(sec);
    }
    // 덜 찬 반 우선
    out.sort((a,b)=> (counts.get(a)-counts.get(b)) || a.localeCompare(b));
    return out;
  }

  function scorePair(sec1, sec2, groupSize) {
    // 균형을 위해: 배정 후 최대 카운트가 작게, 총합도 작게
    const c1 = (counts.get(sec1) ?? 0) + groupSize;
    const c2 = (counts.get(sec2) ?? 0) + groupSize;
    const maxC = Math.max(c1, c2);
    const sumC = c1 + c2;
    return [maxC, sumC, sec1, sec2]; // lexicographic 최소
  }

  // 그룹(조합) 단위로 묶기: 학생반+관+탐구1+탐구2
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

  // 큰 그룹부터 배정(분할 최소화에 유리)
  const groupEntries = [...groups.entries()]
    .map(([key, members]) => ({ key, members }))
    .sort((a,b)=> b.members.length - a.members.length);

  // 결과 행 저장
  const results = [];

  // 그룹 배정 기록(같은 조합은 항상 같은 pair 시도)
  const groupPairMap = new Map(); // key -> {sec1, sec2}

  // 그룹 배정 함수: 가능한 경우 그룹 전체 동일 pair, 아니면 분할(최소 분할)
  function assignGroup({key, members}) {
    const [home, building, subj1, subj2] = key.split("||");
    const groupSize = members.length;

    const blockedBase = new Set();
    const mandSlots = mandatorySlotsByHome.get(home);
    if (mandSlots) for (const x of mandSlots) blockedBase.add(x);

    let warnAll = false;
    const notesBase = [];
    if (!mandSlots) {
      warnAll = true;
      notesBase.push(`필수시간표 없음(반=${home})`);
    }

    // 1) 이미 같은 조합으로 성공했던 pair가 있으면 우선 시도
    const preferred = groupPairMap.get(key);
    const tryPreferred = (pair, size) => {
      if (!pair) return null;
      const {sec1, sec2} = pair;
      if (!sec1 || !sec2) return null;
      if (!sections.has(sec1) || !sections.has(sec2)) return null;
      if (!canFit(sec1, size) || !canFit(sec2, size)) return null;
      if (hasConflict(sections.get(sec1).slots, blockedBase)) return null;
      const blocked = new Set(blockedBase);
      for (const x of sections.get(sec1).slots) blocked.add(x);
      if (hasConflict(sections.get(sec2).slots, blocked)) return null;
      return {sec1, sec2};
    };

    const okPreferred = tryPreferred(preferred, groupSize);
    if (okPreferred) {
      counts.set(okPreferred.sec1, (counts.get(okPreferred.sec1)??0) + groupSize);
      counts.set(okPreferred.sec2, (counts.get(okPreferred.sec2)??0) + groupSize);
      return [{ size: groupSize, sec1: okPreferred.sec1, sec2: okPreferred.sec2, warn: warnAll, notes: notesBase }];
    }

    // 2) 전체 그룹을 수용 가능한 pair 탐색
    const cand1 = getCandidates(subj1, building, blockedBase, groupSize);
    let best = null; // {sec1, sec2, score}
    for (const sec1 of cand1) {
      const blocked = new Set(blockedBase);
      for (const x of sections.get(sec1).slots) blocked.add(x);

      const cand2 = getCandidates(subj2, building, blocked, groupSize);
      for (const sec2 of cand2) {
        // 탐구1/2가 같은 반이 되는 경우는 일반적으로 불가능(시간/과목)하지만 일단 허용 X로 막고 싶으면 아래 조건 추가
        // if (sec1 === sec2) continue;

        const sc = scorePair(sec1, sec2, groupSize);
        if (!best || sc < best.score) best = { sec1, sec2, score: sc };
      }
    }

    if (best) {
      groupPairMap.set(key, {sec1: best.sec1, sec2: best.sec2});
      counts.set(best.sec1, (counts.get(best.sec1)??0) + groupSize);
      counts.set(best.sec2, (counts.get(best.sec2)??0) + groupSize);
      return [{ size: groupSize, sec1: best.sec1, sec2: best.sec2, warn: warnAll, notes: notesBase }];
    }

    // 3) 전체 수용 불가 → 분할(최소 분할)
    // 전략: 가능한 최대 size를 먼저 배정하고, 나머지 재귀적으로 처리
    // groupSize가 크면 brute-force가 커지니, 최대 size는 remainingCap 기반으로 빠르게 계산
    let remaining = groupSize;
    const chunks = [];
    let split = false;

    while (remaining > 0) {
      // 현재 remaining에서 들어갈 수 있는 최대 chunk size 찾기
      let found = null;

      // 최대치부터 줄여가며 찾음(분할 최소화를 우선)
      for (let chunkSize = remaining; chunkSize >= 1; chunkSize--) {
        // preferred 재시도
        const okPref = tryPreferred(preferred, chunkSize);
        if (okPref) {
          found = { size: chunkSize, sec1: okPref.sec1, sec2: okPref.sec2 };
          break;
        }

        const c1 = getCandidates(subj1, building, blockedBase, chunkSize);
        let bestLocal = null;
        for (const sec1 of c1) {
          const blocked = new Set(blockedBase);
          for (const x of sections.get(sec1).slots) blocked.add(x);

          const c2 = getCandidates(subj2, building, blocked, chunkSize);
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
        // 남은 학생들은 배정 불가
        split = true;
        chunks.push({
          size: remaining,
          sec1: "",
          sec2: "",
          warn: true,
          notes: [...notesBase, "조합 단위 배정 실패(관/정원/충돌/개설없음)"]
        });
        remaining = 0;
        break;
      }

      if (found.size !== groupSize) split = true;

      // 카운트 반영
      counts.set(found.sec1, (counts.get(found.sec1)??0) + found.size);
      counts.set(found.sec2, (counts.get(found.sec2)??0) + found.size);

      chunks.push({
        size: found.size,
        sec1: found.sec1,
        sec2: found.sec2,
        warn: warnAll || split,
        notes: split ? [...notesBase, "⚠️ 같은 조합 그룹이 정원/제약으로 분할 배정됨"] : notesBase
      });

      remaining -= found.size;
    }

    // split이 없고 모두 배정된 경우에만 groupPairMap 저장(다음 그룹 재사용)
    const allAssignedOneChunk = chunks.length === 1 && chunks[0].sec1 && chunks[0].sec2 && chunks[0].size === groupSize;
    if (allAssignedOneChunk) {
      groupPairMap.set(key, {sec1: chunks[0].sec1, sec2: chunks[0].sec2});
    }
    return chunks;
  }

  // 실제 배정 수행
  for (const g of groupEntries) {
    const chunks = assignGroup(g);

    // chunks에 따라 멤버를 앞에서부터 잘라 배정(분할 시)
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
          일련번호: id,
          관: building,
          학생반: home,
          탐구1: subj1,
          탐구1_배정반: ch.sec1,
          탐구1_시간표: ch.sec1 ? sections.get(ch.sec1).detail : "",
          탐구2: subj2,
          탐구2_배정반: ch.sec2,
          탐구2_시간표: ch.sec2 ? sections.get(ch.sec2).detail : "",
          경고: warn ? "⚠️" : "",
          비고: notes.filter(Boolean).join(" | "),
          "필수시간표(참고)": mandatoryDetailByHome.get(home) ?? "",
          "조합그룹키(참고)": g.key
        });
      }
    }
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
