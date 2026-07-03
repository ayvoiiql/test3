/* ============================================================
   WorkMind — app.js  (완전 동작 버전)
   Vanilla ES6+ / LocalStorage / SheetJS(Excel)
   ------------------------------------------------------------
   기능
   ① AI 업무 목표 분해 (현재: 로컬 생성기 / 실제 API 훅 포함)
   ② 업무 대시보드 (스트릭·히트맵·진도율·주간추이)
   ③ 스마트 리마인더 (Notification API)
   ④ 다중 목표 관리 (CRUD + 아카이브)
   ⑤ 오늘의 할 일 (통합 체크리스트 + 이월)
   + LocalStorage 자동 저장 / Excel 백업·복원
============================================================ */

(function () {
  'use strict';

  /* ============================================================
     0. 상수 / 유틸
  ============================================================ */

  const STORAGE_KEY = 'studymind:v1';

  const SCREENS = [
    'screen-welcome', 'screen-today', 'screen-goals', 'screen-goal-new',
    'screen-goal-detail', 'screen-goal-edit', 'screen-dashboard', 'screen-settings',
  ];

  const NAV_MAP = {
    'screen-today': 'nav-today',
    'screen-goals': 'nav-goals',
    'screen-goal-new': 'nav-goals',
    'screen-goal-detail': 'nav-goals',
    'screen-goal-edit': 'nav-goals',
    'screen-dashboard': 'nav-dashboard',
    'screen-settings': 'nav-settings',
  };

  // 목표 식별 색상 (기능 ④/⑤ 색상 태그)
  const GOAL_COLORS = ['#7c3aed', '#0066cc', '#0d9488', '#ea580c', '#db2777', '#ca8a04'];

  const WEEKDAY_KO = ['일', '월', '화', '수', '목', '금', '토'];

  const $ = (sel, root) => (root || document).querySelector(sel);
  const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));
  const el = (id) => document.getElementById(id);

  // XSS 방지: 사용자 입력을 텍스트로만 삽입 (PRD 4-2)
  function esc(str) {
    const d = document.createElement('div');
    d.textContent = str == null ? '' : String(str);
    return d.innerHTML;
  }

  function uid() {
    return 'id-' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
  }

  // 날짜 헬퍼 (로컬 타임존 기준)
  function toISO(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }
  function parseISO(s) { return new Date(s + 'T00:00:00'); }
  function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
  function todayISO() { return toISO(new Date()); }
  function diffDays(aISO, bISO) {
    return Math.round((parseISO(aISO) - parseISO(bISO)) / 86400000);
  }
  function fmtKDate(d) {
    return `${d.getFullYear()}.${String(d.getMonth() + 1).padStart(2, '0')}.${String(d.getDate()).padStart(2, '0')} (${WEEKDAY_KO[d.getDay()]})`;
  }
  function fmtShort(iso) {
    const d = parseISO(iso);
    return `${d.getMonth() + 1}/${d.getDate()} (${WEEKDAY_KO[d.getDay()]})`;
  }
  function fmtMinutes(min) {
    const h = Math.floor(min / 60), m = min % 60;
    if (h && m) return `${h}시간 ${m}분`;
    if (h) return `${h}시간`;
    return `${m}분`;
  }
  // 이번 주 월요일
  function weekStartISO(iso) {
    const d = parseISO(iso);
    const day = (d.getDay() + 6) % 7; // 월=0
    return toISO(addDays(d, -day));
  }

  /* ============================================================
     1. 상태 (State) + 영속화
  ============================================================ */

  const defaultState = () => ({
    goals: [],   // {id,name,category,level,startDate,endDate,dailyTime,memo,color,archived,createdAt}
    tasks: [],   // {id,goalId,date,text,minutes,done}
    settings: {
      studyAlarm: false, alarmTime: '19:00', alarmDays: [1, 2, 3, 4, 5],
      remindAlarm: false, remindTime: '22:00',
    },
  });

  let state = defaultState();

  function load() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) state = Object.assign(defaultState(), JSON.parse(raw));
    } catch (e) { console.warn('로드 실패, 초기화', e); state = defaultState(); }
  }
  function save() {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); }
    catch (e) { console.error('저장 실패', e); }
  }

  /* ============================================================
     2. 셀렉터 / 계산 (파생 데이터)
  ============================================================ */

  const activeGoals = () => state.goals.filter(g => !g.archived);
  const goalById = (id) => state.goals.find(g => g.id === id);
  const tasksOfGoal = (id) => state.tasks.filter(t => t.goalId === id);

  // 목표 진도율(%)
  function goalProgress(id) {
    const ts = tasksOfGoal(id);
    if (!ts.length) return 0;
    return Math.round(ts.filter(t => t.done).length / ts.length * 100);
  }
  // 목표 D-Day (양수=남음)
  function goalDday(g) { return diffDays(g.endDate, todayISO()); }

  // 오늘 뷰 과제 = 오늘 날짜 + (과거 미완료 = 이월)
  function todayTasks() {
    const t = todayISO();
    return state.tasks.filter(x => x.date === t || (x.date < t && !x.done))
      .map(x => Object.assign({ carried: x.date < t }, x));
  }

  // 투데이 화면에서 선택된 날짜의 과제 (오늘이면 이월 포함, 그 외 날짜는 해당일 과제만)
  function tasksForDate(iso) {
    if (iso === todayISO()) return todayTasks();
    return state.tasks.filter(x => x.date === iso).map(x => Object.assign({ carried: false }, x));
  }

  // 특정 날짜 통계
  function dayStats(iso) {
    const ts = state.tasks.filter(t => t.date === iso);
    return { total: ts.length, done: ts.filter(t => t.done).length };
  }

  // 연속 실행일(스트릭): 오늘(또는 어제)부터 뒤로 "완료 과제 ≥1"인 날 연속
  function calcStreak() {
    let streak = 0;
    let cur = new Date();
    // 오늘 아직 완료 0이면 어제부터 카운트 (스트릭 유지)
    if (dayStats(toISO(cur)).done === 0) cur = addDays(cur, -1);
    for (let i = 0; i < 3650; i++) {
      if (dayStats(toISO(cur)).done > 0) { streak++; cur = addDays(cur, -1); }
      else break;
    }
    return streak;
  }
  // 총 실행일 (완료 과제가 하나라도 있는 날 수)
  function totalStudyDays() {
    const set = new Set(state.tasks.filter(t => t.done).map(t => t.date));
    return set.size;
  }

  /* ============================================================
     3. AI 업무 목표 분해 (기능 ①)
     - 현재: 로컬 휴리스틱 생성기 (오프라인 동작)
     - 실제 배포: callRealAI() 를 Vercel 서버리스로 연결
  ============================================================ */

  const TASK_TEMPLATES = {
    language: [
      ['주간 업무 보고서 작성', 30], ['이메일 회신 {q}건 처리', 20],
      ['회의 안건 정리', 25], ['업무 매뉴얼 Chapter {n} 검토', 20],
      ['피드백 반영 수정', 15], ['거래처 미팅 준비', 20], ['업무 자료 {q}건 검토', 25],
    ],
    coding: [
      ['사이드 프로젝트 기능 개발', 60], ['{topic} 리서치 및 정리', 20],
      ['프로토타입 제작', 45], ['프로젝트 회고 작성', 30],
      ['버그 수정 및 리팩터링', 25], ['{topic} 관련 레퍼런스 조사', 30],
    ],
    certificate: [
      ['자격증 기출문제 {q}개 풀기', 40], ['이력서/포트폴리오 업데이트', 30],
      ['업계 트렌드 아티클 {q}편 읽기', 20], ['핵심 이론 {q}개 정리', 25], ['모의 면접 준비', 50],
    ],
    etc: [
      ['업무 자료 {q}p 검토', 30], ['핵심 내용 정리', 20],
      ['실행 항목 {q}개 처리', 30], ['정리 노트 작성', 20], ['체크리스트 {q}항목 점검', 15],
    ],
  };
  const CODING_TOPICS = ['기획', '디자인', 'UX 리서치', '마케팅', '데이터 분석', '브랜딩', '콘텐츠 제작', '수익화 전략'];

  function fill(tmpl, seed) {
    return tmpl
      .replace('{n}', (seed % 8) + 1)
      .replace('{q}', 10 + (seed % 4) * 5)
      .replace('{topic}', CODING_TOPICS[seed % CODING_TOPICS.length]);
  }

  // 하루 가용시간 → 하루 과제 개수 (3~5개)
  function tasksPerDay(min) {
    if (min <= 30) return 3;
    if (min <= 60) return 3;
    if (min <= 120) return 4;
    return 5;
  }

  // 로컬 생성: 목표 기간 동안 날짜별 과제 배열 반환
  function generateTasksLocal(goal, fromISO) {
    const cat = TASK_TEMPLATES[goal.category] ? goal.category : 'etc';
    const templates = TASK_TEMPLATES[cat];
    const start = fromISO && fromISO > goal.startDate ? fromISO : goal.startDate;
    // 안전: 시작일이 오늘보다 과거면 오늘부터 생성
    const realStart = start < todayISO() ? todayISO() : start;
    const spanDays = Math.max(1, Math.min(90, diffDays(goal.endDate, realStart) + 1));
    const perDay = tasksPerDay(goal.dailyTime);
    const out = [];
    let seed = 0;
    for (let i = 0; i < spanDays; i++) {
      const dISO = toISO(addDays(parseISO(realStart), i));
      if (dISO > goal.endDate) break;
      for (let j = 0; j < perDay; j++) {
        const [tmpl, min] = templates[(i + j) % templates.length];
        out.push({
          id: uid(), goalId: goal.id, date: dISO,
          text: fill(tmpl, seed++), minutes: min, done: false,
        });
      }
    }
    return out;
  }

  // 실제 AI API 훅 (배포 시 사용). 실패하면 로컬 생성기로 폴백.
  async function callRealAI(goal) {
    // 예시: Vercel 서버리스 프록시 호출
    // const res = await fetch('/api/generate', {
    //   method: 'POST', headers: { 'Content-Type': 'application/json' },
    //   body: JSON.stringify(goal),
    // });
    // if (!res.ok) throw new Error('AI 응답 실패');
    // const { tasks } = await res.json();
    // return tasks.map(t => ({ id: uid(), goalId: goal.id, ...t, done: false }));
    throw new Error('AI 미연결 — 로컬 생성기 사용');
  }

  const AI_TIPS = [
    '작은 과제를 꾸준히 완료하면 큰 목표도 이룰 수 있어요!',
    '오늘 할 일에 집중하세요. 계획은 저희가 짤게요.',
    '연속 실행일을 늘리면 습관이 됩니다 🔥',
    '완료 체크의 쾌감, 오늘도 느껴보세요 ✅',
  ];

  // 목표에 대한 계획 생성 (모달 표시 → 생성 → 저장)
  async function generatePlan(goal, replaceExisting) {
    const dlg = el('modal-ai-loading');
    const prog = el('ai-loading-progress');
    const tip = el('ai-loading-tip');
    const errBox = el('ai-loading-error');
    errBox.hidden = true;
    tip.textContent = '💡 팁: ' + AI_TIPS[Math.floor(Math.random() * AI_TIPS.length)];
    if (dlg.showModal && !dlg.open) dlg.showModal();

    // 가짜 진행바 (0→90%)
    prog.value = 0;
    const timer = setInterval(() => { prog.value = Math.min(90, prog.value + 8); }, 150);

    let tasks;
    try {
      tasks = await callRealAI(goal);       // 실제 AI 시도
    } catch (e) {
      await new Promise(r => setTimeout(r, 1200)); // 로딩 체감용
      tasks = generateTasksLocal(goal, replaceExisting ? todayISO() : null);
    }

    clearInterval(timer);
    prog.value = 100;

    if (replaceExisting) {
      // 재생성: 오늘 이후 미완료 과제만 교체 (완료 기록 보존)
      state.tasks = state.tasks.filter(t =>
        t.goalId !== goal.id || t.done || t.date < todayISO());
    }
    state.tasks.push(...tasks);
    save();

    setTimeout(() => { if (dlg.open) dlg.close(); }, 250);
    return tasks;
  }

  /* ============================================================
     4. 라우팅
  ============================================================ */

  function normalizeHash() {
    const id = (location.hash || '').replace(/^#/, '');
    if (SCREENS.includes(id)) return id;
    return activeGoals().length ? 'screen-today' : 'screen-welcome';
  }

  let currentGoalId = null; // 상세/수정 대상
  let selectedDate = todayISO(); // 투데이 화면에서 보고 있는 날짜

  function showScreen(screenId) {
    SCREENS.forEach(id => {
      const s = el(id);
      if (s) s.classList.toggle('is-active', id === screenId);
    });
    const navId = NAV_MAP[screenId];
    $$('#global-nav a').forEach(a => a.classList.toggle('is-active', a.id === navId));
    window.scrollTo(0, 0);
    renderScreen(screenId);
  }

  function handleRoute() { showScreen(normalizeHash()); }

  function navigateTo(screenId) {
    if (location.hash === '#' + screenId) handleRoute();
    else location.hash = screenId;
  }

  // 화면 진입 시 필요한 렌더 트리거
  function renderScreen(id) {
    switch (id) {
      case 'screen-today': renderToday(); break;
      case 'screen-goals': renderGoals(); break;
      case 'screen-goal-detail': renderGoalDetail(); break;
      case 'screen-goal-edit': renderGoalEditForm(); break;
      case 'screen-dashboard': renderDashboard(); break;
      case 'screen-goal-new': resetNewForm(); break;
    }
  }

  /* ============================================================
     5. 렌더링 — 투데이 뷰 (기능 ⑤)
  ============================================================ */

  function renderToday() {
    const isToday = selectedDate === todayISO();
    el('today-date').textContent = fmtKDate(parseISO(selectedDate)) + (isToday ? ' · 오늘' : '');
    el('today-date').setAttribute('datetime', selectedDate);
    el('today-date-picker').value = selectedDate;
    el('today-date-today-btn').hidden = isToday;

    const hasGoals = activeGoals().length > 0;
    el('today-empty-state').hidden = hasGoals;
    el('today-summary').style.display = hasGoals ? '' : 'none';
    el('today-tasklist').style.display = hasGoals ? '' : 'none';
    el('today-footer-info').style.display = hasGoals ? '' : 'none';
    if (!hasGoals) return;

    const tasks = tasksForDate(selectedDate);
    const done = tasks.filter(t => t.done).length;
    const total = tasks.length;
    const pct = total ? Math.round(done / total * 100) : 0;
    const streak = calcStreak();

    // 요약 카드
    el('today-summary').innerHTML = `
      <div id="today-summary-streak">🔥 <span class="value">${streak}</span>일 연속</div>
      <div id="today-summary-count">✅ <span class="value">${done}/${total}</span> 완료</div>
      <progress id="today-summary-progress" value="${pct}" max="100">${pct}%</progress>
      <span id="today-summary-percent">${pct}%</span>`;

    // 목표별 그룹
    const byGoal = {};
    tasks.forEach(t => { (byGoal[t.goalId] = byGoal[t.goalId] || []).push(t); });
    const wrap = el('today-task-list');
    wrap.innerHTML = '';
    if (!tasks.length) {
      wrap.innerHTML = `<p class="today-empty-day">이 날짜에는 할 일이 없어요.</p>`;
    }
    Object.keys(byGoal).forEach(gid => {
      const g = goalById(gid);
      if (!g) return;
      const group = document.createElement('div');
      group.className = 'today-goal-group';
      group.innerHTML = `
        <h3 class="today-goal-name"><span class="goal-dot" style="background:${g.color}"></span>${esc(g.name)}</h3>
        <ul class="today-task-items">
          ${byGoal[gid].map(t => taskLi(t, 'today')).join('')}
        </ul>`;
      wrap.appendChild(group);
    });

    // 푸터: 예상 시간 / 이월
    const totalMin = tasks.reduce((s, t) => s + (t.minutes || 0), 0);
    const carried = tasks.filter(t => t.carried).length;
    el('today-footer-info').innerHTML = `
      <p id="today-estimated-time">⏰ ${isToday ? '오늘' : '이 날'} 예상 소요 시간: ${fmtMinutes(totalMin)}</p>
      ${carried ? `<p id="today-carryover-notice">※ 어제까지 미완료 과제 ${carried}건 이월됨 ⚠️</p>` : ''}`;
  }

  // 과제 <li> HTML (context: 'today' | 'detail')
  function taskLi(t, ctx) {
    const cls = ctx === 'today' ? 'today-task-item' : 'detail-task-item';
    return `<li class="${cls}" data-task-id="${t.id}">
      <input type="checkbox" ${t.done ? 'checked' : ''} data-task-check="${t.id}">
      <span class="task-text">${esc(t.text)}${t.carried ? ' <span class="carry-badge">이월</span>' : ''}</span>
      <span class="task-time">${t.minutes}분</span>
    </li>`;
  }

  /* ============================================================
     6. 렌더링 — 목표 목록 (기능 ④)
  ============================================================ */

  function renderGoals() {
    const active = activeGoals();
    const done = state.goals.filter(g => g.archived);

    el('goals-active-title').textContent = `진행 중 (${active.length})`;
    el('goals-completed-title').textContent = `완료됨 (${done.length})`;
    el('goals-completed').style.display = done.length ? '' : 'none';

    const t = todayISO();
    el('goals-active-list').innerHTML = active.length ? active.map(g => {
      const remaining = state.tasks.filter(x => x.goalId === g.id && x.date === t && !x.done).length;
      const dd = goalDday(g);
      const pct = goalProgress(g.id);
      return `<li class="goal-card" data-goal-id="${g.id}" tabindex="0">
        <h3 class="goal-card-name"><span class="goal-dot" style="background:${g.color}"></span>${esc(g.name)}</h3>
        <span class="goal-card-dday">${dd >= 0 ? 'D-' + dd : 'D+' + (-dd)}</span>
        <progress class="goal-card-progress" value="${pct}" max="100">${pct}%</progress>
        <span class="goal-card-percent">${pct}%</span>
        <p class="goal-card-remaining">오늘 과제: ${remaining}개 남음</p>
      </li>`;
    }).join('') : `<li class="empty-hint">아직 목표가 없어요. [+ 추가]로 시작하세요.</li>`;

    el('goals-completed-list').innerHTML = done.map(g =>
      `<li class="goal-card goal-card-done" data-goal-id="${g.id}" tabindex="0">
        <h3 class="goal-card-name">✅ ${esc(g.name)}</h3>
        <span class="goal-card-percent">${goalProgress(g.id)}%</span>
      </li>`).join('');
  }

  /* ============================================================
     7. 렌더링 — 목표 상세 (기능 ①④)
  ============================================================ */

  let detailFilter = 'today';

  function renderGoalDetail() {
    const g = goalById(currentGoalId);
    if (!g) { navigateTo('screen-goals'); return; }

    el('goal-detail-title').innerHTML =
      `<span class="goal-dot" style="background:${g.color}"></span>${esc(g.name)}`;

    const dd = goalDday(g);
    const pct = goalProgress(g.id);
    el('goal-detail-info').innerHTML = `
      <p>📅 ${g.startDate} ~ ${g.endDate} <span id="goal-detail-dday">${dd >= 0 ? 'D-' + dd : 'D+' + (-dd)}</span></p>
      <p>📈 진도율 ${pct}%</p>
      <progress value="${pct}" max="100">${pct}%</progress>`;

    // 탭 활성
    $$('#goal-detail-tabs .goal-detail-tab').forEach(b =>
      b.classList.toggle('is-active', b.dataset.filter === detailFilter));

    // 필터링된 과제
    const t = todayISO();
    const wStart = weekStartISO(t), wEnd = toISO(addDays(parseISO(wStart), 6));
    let list = tasksOfGoal(g.id);
    if (detailFilter === 'today') list = list.filter(x => x.date === t || (x.date < t && !x.done));
    else if (detailFilter === 'week') list = list.filter(x => x.date >= wStart && x.date <= wEnd);
    list.sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : 0);

    // 날짜별 그룹
    const byDate = {};
    list.forEach(x => { (byDate[x.date] = byDate[x.date] || []).push(x); });
    const dates = Object.keys(byDate).sort();

    const cont = el('goal-detail-tasks');
    cont.innerHTML = dates.length ? dates.map(d => {
      let label = fmtShort(d);
      if (d === t) label += ' 오늘';
      else if (d === toISO(addDays(new Date(), 1))) label += ' 내일';
      return `<div class="goal-detail-day-group" data-date="${d}">
        <h2 class="goal-detail-day-title">${label}</h2>
        <ul class="goal-detail-day-tasks">${byDate[d].map(x => taskLi(x, 'detail')).join('')}</ul>
        ${d === t ? `<button type="button" class="goal-detail-btn-add-task" data-add-date="${d}">+ 과제 직접 추가</button>` : ''}
      </div>`;
    }).join('') : `<p class="empty-hint">이 필터에 해당하는 과제가 없어요.</p>`;
  }

  /* ============================================================
     8. 렌더링 — 목표 수정 폼
  ============================================================ */

  function renderGoalEditForm() {
    const g = goalById(currentGoalId);
    if (!g) { navigateTo('screen-goals'); return; }
    el('edit-goal-name').value = g.name;
    el('edit-goal-level').value = g.level || '';
    el('edit-goal-start').value = g.startDate;
    el('edit-goal-end').value = g.endDate;
    el('edit-goal-memo').value = g.memo || '';
    selectChip('#edit-category-chips', 'category', g.category);
    selectChip('#edit-dailytime-chips', 'time', String(g.dailyTime));
  }

  /* ============================================================
     9. 렌더링 — 대시보드 (기능 ②)
  ============================================================ */

  function renderDashboard() {
    const hasData = state.tasks.some(t => t.done);
    el('dashboard-empty-state').hidden = hasData;

    // 핵심 지표
    const t = todayISO();
    const ds = dayStats(t);
    el('dashboard-metrics').innerHTML = `
      <div class="dashboard-metric" id="metric-streak">🔥 연속<span class="value">${calcStreak()}일</span></div>
      <div class="dashboard-metric" id="metric-today">✅ 오늘<span class="value">${ds.done}/${ds.total} 완료</span></div>
      <div class="dashboard-metric" id="metric-total">📅 총<span class="value">${totalStudyDays()}일</span></div>`;

    renderHeatmap();
    renderGoalProgressList();
    renderWeeklyTrend();
  }

  // 히트맵: 최근 8주(월~일) 격자
  function renderHeatmap() {
    const grid = el('dashboard-heatmap-grid');
    const weeks = 8;
    const start = parseISO(weekStartISO(todayISO()));
    start.setDate(start.getDate() - (weeks - 1) * 7);

    let cells = '';
    for (let w = 0; w < weeks; w++) {
      for (let d = 0; d < 7; d++) {
        const iso = toISO(addDays(start, w * 7 + d));
        const { total, done } = dayStats(iso);
        let lvl = 0;
        if (total > 0 && done > 0) lvl = (done / total) <= 0.5 ? 1 : 2;
        const future = iso > todayISO();
        cells += `<span class="heat-cell heat-${lvl}${future ? ' heat-future' : ''}" title="${iso} · ${done}/${total}"></span>`;
      }
    }
    grid.innerHTML = `<div class="heat-weekdays">${WEEKDAY_KO.map(w => `<span>${w}</span>`).join('')}</div>
      <div class="heat-grid">${cells}</div>`;
  }

  function renderGoalProgressList() {
    const list = el('dashboard-goal-progress-list');
    const gs = activeGoals();
    list.innerHTML = gs.length ? gs.map(g => {
      const pct = goalProgress(g.id);
      return `<li class="dashboard-goal-progress-item">
        <span class="goal-name"><span class="goal-dot" style="background:${g.color}"></span>${esc(g.name)}</span>
        <span class="percent">${pct}%</span>
        <progress value="${pct}" max="100">${pct}%</progress>
      </li>`;
    }).join('') : `<li class="empty-hint">목표가 없어요.</li>`;
  }

  // 주간 완료율 추이: 최근 6주 (SVG 막대)
  function renderWeeklyTrend() {
    const weeks = 6;
    const data = [];
    const thisWeekStart = parseISO(weekStartISO(todayISO()));
    for (let i = weeks - 1; i >= 0; i--) {
      const ws = addDays(thisWeekStart, -i * 7);
      const wsISO = toISO(ws), weISO = toISO(addDays(ws, 6));
      const ts = state.tasks.filter(x => x.date >= wsISO && x.date <= weISO);
      const rate = ts.length ? Math.round(ts.filter(x => x.done).length / ts.length * 100) : 0;
      data.push({ label: `${weeks - i - (weeks - 1) + i}주`, rate, ws });
    }
    // 라벨을 1주..6주로 단순화
    data.forEach((d, i) => d.label = `${i + 1}주`);

    const W = 320, H = 160, pad = 24, bw = (W - pad * 2) / weeks;
    let bars = '';
    data.forEach((d, i) => {
      const bh = (H - pad * 2) * (d.rate / 100);
      const x = pad + i * bw + bw * 0.2;
      const y = H - pad - bh;
      bars += `<rect x="${x}" y="${y}" width="${bw * 0.6}" height="${bh}" rx="3" fill="var(--color-primary)"></rect>
        <text x="${x + bw * 0.3}" y="${H - pad + 14}" text-anchor="middle" font-size="10" fill="#7a7a7a">${d.label}</text>
        <text x="${x + bw * 0.3}" y="${y - 4}" text-anchor="middle" font-size="9" fill="#7a7a7a">${d.rate}%</text>`;
    });
    el('dashboard-weekly-chart').innerHTML =
      `<svg viewBox="0 0 ${W} ${H}" width="100%" role="img" aria-label="주간 완료율 추이">
        <line x1="${pad}" y1="${H - pad}" x2="${W - pad}" y2="${H - pad}" stroke="#e0e0e0"></line>
        ${bars}
      </svg>`;
  }

  /* ============================================================
     10. 이벤트 — 과제 체크 (위임)
  ============================================================ */

  function onTaskToggle(e) {
    const cb = e.target.closest('[data-task-check]');
    if (!cb) return;
    const id = cb.getAttribute('data-task-check');
    const task = state.tasks.find(t => t.id === id);
    if (!task) return;
    task.done = cb.checked;
    save();

    // 현재 화면 부분 갱신
    const li = cb.closest('li');
    if (li) li.querySelector('.task-text').classList.toggle('is-done', cb.checked);

    // 투데이 요약/진도 즉시 갱신
    if (el('screen-today').classList.contains('is-active')) {
      const tasks = tasksForDate(selectedDate);
      const done = tasks.filter(t => t.done).length, total = tasks.length;
      const pct = total ? Math.round(done / total * 100) : 0;
      const sp = el('today-summary-progress'); if (sp) sp.value = pct;
      const spc = el('today-summary-percent'); if (spc) spc.textContent = pct + '%';
      const cnt = $('#today-summary-count .value'); if (cnt) cnt.textContent = `${done}/${total}`;
      const stk = $('#today-summary-streak .value'); if (stk) stk.textContent = calcStreak();

      // 전체 완료 시 축하 모달 (오늘 날짜를 볼 때만)
      if (selectedDate === todayISO() && total > 0 && done === total && cb.checked) showCelebration(done);
    }
    if (el('screen-goal-detail').classList.contains('is-active')) renderGoalDetail();
  }

  function showCelebration(count) {
    const streak = calcStreak();
    el('celebration-streak').textContent = `🔥 ${streak}일 연속 실행 중`;
    el('celebration-count').textContent = `✅ 오늘 ${count}개 과제 완료`;
    const dlg = el('modal-celebration');
    if (dlg.showModal && !dlg.open) dlg.showModal();
  }

  /* ============================================================
     11. 폼 — 칩 선택 헬퍼
  ============================================================ */

  function wireChipGroup(sel, attr) {
    const group = $(sel);
    if (!group) return;
    group.addEventListener('click', e => {
      const chip = e.target.closest('.chip');
      if (!chip) return;
      $$('.chip', group).forEach(c => c.classList.remove('is-selected'));
      chip.classList.add('is-selected');
    });
  }
  function selectedChip(sel, attr) {
    const c = $(sel + ' .chip.is-selected');
    return c ? c.dataset[attr] : null;
  }
  function selectChip(sel, attr, val) {
    $$(sel + ' .chip').forEach(c =>
      c.classList.toggle('is-selected', c.dataset[attr] === val));
  }

  function resetNewForm() {
    el('goal-new-form').reset();
    el('goal-new-error').hidden = true;
    $$('#goal-category-chips .chip, #goal-dailytime-chips .chip').forEach(c => c.classList.remove('is-selected'));
    // 기본 날짜: 오늘 ~ +30일
    el('input-goal-start').value = todayISO();
    el('input-goal-end').value = toISO(addDays(new Date(), 30));
  }

  /* ============================================================
     12. 폼 제출 — 목표 생성 / 수정 / 삭제
  ============================================================ */

  async function onCreateGoal(e) {
    e.preventDefault();
    const name = el('input-goal-name').value.trim();
    const start = el('input-goal-start').value;
    const end = el('input-goal-end').value;
    const time = selectedChip('#goal-dailytime-chips', 'time');
    const err = el('goal-new-error');

    if (!name || !start || !end || !time) {
      err.textContent = '필수 항목(목표·기간·가용시간)을 모두 입력해주세요.';
      err.hidden = false; return;
    }
    if (end < start) { err.textContent = '종료일이 시작일보다 빠를 수 없어요.'; err.hidden = false; return; }
    err.hidden = true;

    const goal = {
      id: uid(), name,
      category: selectedChip('#goal-category-chips', 'category') || 'etc',
      level: el('input-goal-level').value.trim(),
      startDate: start, endDate: end, dailyTime: Number(time),
      memo: el('input-goal-memo').value.trim(),
      color: GOAL_COLORS[state.goals.length % GOAL_COLORS.length],
      archived: false, createdAt: todayISO(),
    };
    state.goals.push(goal);
    save();

    await generatePlan(goal, false);
    currentGoalId = goal.id;
    detailFilter = 'today';
    navigateTo('screen-goal-detail');
  }

  function readEditForm(g) {
    g.name = el('edit-goal-name').value.trim() || g.name;
    g.level = el('edit-goal-level').value.trim();
    g.startDate = el('edit-goal-start').value || g.startDate;
    g.endDate = el('edit-goal-end').value || g.endDate;
    g.memo = el('edit-goal-memo').value.trim();
    g.category = selectedChip('#edit-category-chips', 'category') || g.category;
    const time = selectedChip('#edit-dailytime-chips', 'time');
    if (time) g.dailyTime = Number(time);
  }

  function onSaveGoal(e) {
    e.preventDefault();
    const g = goalById(currentGoalId); if (!g) return;
    readEditForm(g); save();
    navigateTo('screen-goal-detail');
  }
  async function onSaveRegenerate() {
    const g = goalById(currentGoalId); if (!g) return;
    readEditForm(g); save();
    await generatePlan(g, true);
    navigateTo('screen-goal-detail');
  }
  function onDeleteGoal() {
    const g = goalById(currentGoalId); if (!g) return;
    el('confirm-delete-message').textContent =
      `"${g.name}" 목표를 삭제하면 관련된 모든 과제와 기록이 함께 사라져요. 되돌릴 수 없어요.`;
    const dlg = el('modal-confirm-delete');
    if (dlg.showModal && !dlg.open) dlg.showModal();
  }
  function onConfirmDeleteGoal() {
    const g = goalById(currentGoalId); if (!g) return;
    state.tasks = state.tasks.filter(t => t.goalId !== g.id);
    state.goals = state.goals.filter(x => x.id !== g.id);
    save();
    el('modal-confirm-delete').close();
    navigateTo('screen-goals');
  }
  async function onRegenerate() {
    const g = goalById(currentGoalId); if (!g) return;
    if (!confirm('남은 기간 기준으로 과제를 다시 생성할까요? (완료한 과제는 유지됩니다)')) return;
    await generatePlan(g, true);
    renderGoalDetail();
  }

  // 과제 직접 추가 (인라인 prompt — MVP)
  function onAddTask(e) {
    const btn = e.target.closest('.goal-detail-btn-add-task');
    if (!btn) return;
    const g = goalById(currentGoalId); if (!g) return;
    const text = prompt('추가할 과제 내용을 입력하세요:');
    if (!text || !text.trim()) return;
    const min = parseInt(prompt('예상 소요 시간(분)?', '20'), 10) || 20;
    state.tasks.push({ id: uid(), goalId: g.id, date: btn.dataset.addDate, text: text.trim(), minutes: min, done: false });
    save();
    renderGoalDetail();
  }

  /* ============================================================
     13. 설정 — 알림 / Excel / 초기화 (기능 ③ + 데이터)
  ============================================================ */

  function loadSettingsUI() {
    const s = state.settings;
    el('toggle-study-alarm').checked = s.studyAlarm;
    el('select-alarm-time').value = s.alarmTime;
    el('toggle-incomplete-remind').checked = s.remindAlarm;
    el('select-remind-time').value = s.remindTime;
    $$('.alarm-day').forEach(cb => cb.checked = s.alarmDays.includes(Number(cb.value)));
  }

  function saveSettingsFromUI() {
    const s = state.settings;
    s.studyAlarm = el('toggle-study-alarm').checked;
    s.alarmTime = el('select-alarm-time').value;
    s.remindAlarm = el('toggle-incomplete-remind').checked;
    s.remindTime = el('select-remind-time').value;
    s.alarmDays = $$('.alarm-day').filter(cb => cb.checked).map(cb => Number(cb.value));
    save();
  }

  async function requestNotify() {
    if (!('Notification' in window)) { alert('이 브라우저는 알림을 지원하지 않아요.'); return false; }
    if (Notification.permission === 'granted') return true;
    const p = await Notification.requestPermission();
    return p === 'granted';
  }

  // 1분마다 알림 시간 체크 (앱 열려있을 때만 — PRD 5-5)
  let lastNotified = '';
  function notifyTick() {
    if (!('Notification' in window) || Notification.permission !== 'granted') return;
    const s = state.settings;
    const now = new Date();
    const hhmm = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

    const stamp = todayISO() + ' ' + hhmm;
    if (stamp === lastNotified) return;

    // 작업 알림
    if (s.studyAlarm && s.alarmDays.includes(now.getDay()) && hhmm === s.alarmTime) {
      const t = todayTasks().find(x => !x.done);
      const g = t ? goalById(t.goalId) : null;
      new Notification('💼 WorkMind', {
        body: t ? `오늘 "${t.text}"${g ? ' ['+g.name+']' : ''} 이(가) 기다리고 있어요!` : '오늘의 할 일을 시작해볼까요?',
      });
      lastNotified = stamp;
    }
    // 미완료 리마인드
    if (s.remindAlarm && hhmm === s.remindTime) {
      const left = todayTasks().filter(x => !x.done).length;
      if (left > 0) {
        new Notification('⏰ 미완료 리마인드', { body: `오늘 남은 과제 ${left}건이 있어요. 마무리해볼까요?` });
        lastNotified = stamp;
      }
    }
  }

  /* ---- Excel 내보내기 (SheetJS) ---- */
  function exportExcel() {
    if (typeof XLSX === 'undefined') { alert('Excel 라이브러리를 불러오지 못했어요.'); return; }
    const wb = XLSX.utils.book_new();

    const goalRows = state.goals.map(g => ({
      목표ID: g.id, 목표명: g.name, 카테고리: g.category, 시작일: g.startDate,
      목표일: g.endDate, 현재수준: g.level || '', 하루시간: g.dailyTime,
      메모: g.memo || '', 색상: g.color, 아카이브: g.archived ? 'Y' : 'N', 진도율: goalProgress(g.id),
    }));
    const taskRows = state.tasks.map(t => ({
      과제ID: t.id, 목표ID: t.goalId, 날짜: t.date, 과제내용: t.text,
      예상시간: t.minutes, 완료여부: t.done ? 'Y' : 'N',
    }));
    // 기록 시트: 날짜별 집계
    const dateSet = Array.from(new Set(state.tasks.map(t => t.date))).sort();
    const recRows = dateSet.map(d => {
      const st = dayStats(d);
      const min = state.tasks.filter(t => t.date === d && t.done).reduce((s, t) => s + t.minutes, 0);
      return { 날짜: d, 완료과제수: st.done, 전체과제수: st.total, 소요시간분: min };
    });

    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(goalRows), '목표');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(taskRows), '과제');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(recRows), '기록');
    // 설정 백업(복원용 JSON) — 히든 시트
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['settings'], [JSON.stringify(state.settings)]]), '_meta');

    const name = `WorkMind_backup_${todayISO().replace(/-/g, '')}.xlsx`;
    XLSX.writeFile(wb, name);
  }

  /* ---- Excel 가져오기 ---- */
  function importExcel(file) {
    if (typeof XLSX === 'undefined') { alert('Excel 라이브러리를 불러오지 못했어요.'); return; }
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: 'array' });
        const goalRows = XLSX.utils.sheet_to_json(wb.Sheets['목표'] || {});
        const taskRows = XLSX.utils.sheet_to_json(wb.Sheets['과제'] || {});
        if (!goalRows.length && !taskRows.length) { alert('유효한 백업 파일이 아니에요.'); return; }
        if (!confirm('기존 데이터를 덮어씁니다. 계속하시겠어요?')) return;

        const ns = defaultState();
        ns.goals = goalRows.map(r => ({
          id: r.목표ID || uid(), name: r.목표명 || '무제', category: r.카테고리 || 'etc',
          startDate: String(r.시작일), endDate: String(r.목표일), level: r.현재수준 || '',
          dailyTime: Number(r.하루시간) || 60, memo: r.메모 || '',
          color: r.색상 || GOAL_COLORS[0], archived: r.아카이브 === 'Y', createdAt: String(r.시작일),
        }));
        ns.tasks = taskRows.map(r => ({
          id: r.과제ID || uid(), goalId: r.목표ID, date: String(r.날짜),
          text: r.과제내용 || '', minutes: Number(r.예상시간) || 0, done: r.완료여부 === 'Y',
        }));
        // 설정 복원
        try {
          const meta = wb.Sheets['_meta'];
          if (meta) {
            const aoa = XLSX.utils.sheet_to_json(meta, { header: 1 });
            const sjson = aoa && aoa[1] && aoa[1][0];
            if (sjson) ns.settings = Object.assign(ns.settings, JSON.parse(sjson));
          }
        } catch (_) {}

        state = ns; save();
        alert('복원 완료!');
        loadSettingsUI();
        navigateTo('screen-today');
      } catch (err) {
        console.error(err); alert('가져오기 실패: 파일을 확인해주세요.');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function resetAll() {
    if (!confirm('정말 모든 데이터를 삭제하시겠어요?')) return;
    const t = prompt('삭제하려면 "studymind" 를 입력하세요:');
    if (t !== 'studymind') { alert('취소되었습니다.'); return; }
    state = defaultState(); save();
    loadSettingsUI();
    location.hash = 'screen-welcome';
    handleRoute();
  }

  /* ============================================================
     14. 이벤트 배선
  ============================================================ */

  function wire() {
    // 라우팅 버튼
    const navBtns = {
      'welcome-btn-create-goal': 'screen-goal-new',
      'today-empty-create-btn': 'screen-goal-new',
      'goals-btn-add': 'screen-goal-new',
      'goal-new-btn-back': 'screen-goals',
      'goal-detail-btn-back': 'screen-goals',
      'goal-edit-btn-back': 'screen-goal-detail',
      'celebration-btn-dashboard': 'screen-dashboard',
    };
    Object.keys(navBtns).forEach(id => {
      const b = el(id); if (!b) return;
      b.addEventListener('click', () => {
        const dlg = b.closest('dialog'); if (dlg && dlg.open) dlg.close();
        navigateTo(navBtns[id]);
      });
    });

    el('goal-detail-btn-edit').addEventListener('click', () => navigateTo('screen-goal-edit'));

    // 칩 그룹
    wireChipGroup('#goal-category-chips', 'category');
    wireChipGroup('#goal-dailytime-chips', 'time');
    wireChipGroup('#edit-category-chips', 'category');
    wireChipGroup('#edit-dailytime-chips', 'time');

    // 폼
    el('goal-new-form').addEventListener('submit', onCreateGoal);
    el('goal-edit-form').addEventListener('submit', onSaveGoal);
    el('goal-edit-btn-save-regenerate').addEventListener('click', onSaveRegenerate);
    el('goal-edit-btn-delete').addEventListener('click', onDeleteGoal);
    el('goal-detail-btn-regenerate').addEventListener('click', onRegenerate);

    // 목표 상세 탭
    el('goal-detail-tabs').addEventListener('click', e => {
      const tab = e.target.closest('.goal-detail-tab');
      if (!tab) return;
      detailFilter = tab.dataset.filter;
      renderGoalDetail();
    });

    // 과제 추가 (위임)
    el('goal-detail-tasks').addEventListener('click', onAddTask);

    // 과제 체크 (위임) — 투데이 + 상세
    el('today-task-list').addEventListener('change', onTaskToggle);
    el('goal-detail-tasks').addEventListener('change', onTaskToggle);

    // 투데이 날짜 이동
    function goToDate(iso) { selectedDate = iso; renderToday(); }
    el('today-date-prev').addEventListener('click', () =>
      goToDate(toISO(addDays(parseISO(selectedDate), -1))));
    el('today-date-next').addEventListener('click', () =>
      goToDate(toISO(addDays(parseISO(selectedDate), 1))));
    el('today-date-today-btn').addEventListener('click', () => goToDate(todayISO()));
    el('today-date').addEventListener('click', () => {
      const picker = el('today-date-picker');
      if (picker.showPicker) picker.showPicker(); else picker.focus();
    });
    el('today-date-picker').addEventListener('change', e => {
      if (e.target.value) goToDate(e.target.value);
    });

    // 목표 카드 클릭 (위임)
    function cardClick(e) {
      const card = e.target.closest('.goal-card'); if (!card) return;
      currentGoalId = card.dataset.goalId; detailFilter = 'today';
      navigateTo('screen-goal-detail');
    }
    el('goals-active-list').addEventListener('click', cardClick);
    el('goals-completed-list').addEventListener('click', cardClick);
    el('dashboard-goal-progress-list').addEventListener('click', cardClick);

    // 모달 닫기
    el('celebration-btn-close').addEventListener('click', () => el('modal-celebration').close());
    el('confirm-delete-btn-cancel').addEventListener('click', () => el('modal-confirm-delete').close());
    el('confirm-delete-btn-confirm').addEventListener('click', onConfirmDeleteGoal);
    $$('dialog').forEach(d => d.addEventListener('click', e => { if (e.target === d) d.close(); }));

    // 설정
    ['toggle-study-alarm', 'select-alarm-time', 'toggle-incomplete-remind', 'select-remind-time']
      .forEach(id => el(id).addEventListener('change', saveSettingsFromUI));
    $$('.alarm-day').forEach(cb => cb.addEventListener('change', saveSettingsFromUI));
    el('toggle-study-alarm').addEventListener('change', async e => {
      if (e.target.checked && !(await requestNotify())) { e.target.checked = false; saveSettingsFromUI(); }
    });
    el('toggle-incomplete-remind').addEventListener('change', async e => {
      if (e.target.checked && !(await requestNotify())) { e.target.checked = false; saveSettingsFromUI(); }
    });

    el('settings-btn-export').addEventListener('click', exportExcel);
    el('settings-btn-import').addEventListener('click', () => el('settings-import-file').click());
    el('settings-import-file').addEventListener('change', e => {
      if (e.target.files[0]) importExcel(e.target.files[0]); e.target.value = '';
    });
    el('welcome-btn-restore').addEventListener('click', () => {
      navigateTo('screen-settings');
      setTimeout(() => el('settings-import-file').click(), 100);
    });
    el('settings-btn-reset').addEventListener('click', resetAll);

    // 라우팅
    window.addEventListener('hashchange', handleRoute);
  }

  /* ============================================================
     15. 초기화
  ============================================================ */

  function init() {
    load();
    wire();
    loadSettingsUI();
    handleRoute();

    // 알림 스케줄러 (앱 열려있는 동안 1분마다)
    setInterval(notifyTick, 60000);

    // 콘솔 디버깅용
    window.WorkMind = { state, save, navigateTo, get: () => state };
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
