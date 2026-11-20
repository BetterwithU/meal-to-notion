/**
 * 학사일정 관리 대시보드 — 웹앱 서버측
 *
 * 배포: clasp deploy (appsscript.json의 webapp 설정으로 "본인만" 접근)
 * 토큰은 PropertiesService에만 두고 브라우저로 내려보내지 않는다.
 */

// ===== 관리자 보안 검증 함수들 =====
function getAdminEmails() {
  var p = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAILS');
  if (!p) return [];
  return p.split(',').map(function(s) { return s.trim().toLowerCase(); }).filter(Boolean);
}

function getAdminPin() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || '';
}

function getCurrentUserEmail() {
  try {
    return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
  } catch (e) {
    return '';
  }
}

function checkAdminAuth(clientPin) {
  var adminEmails = getAdminEmails();
  var configuredPin = getAdminPin();
  var currentEmail = getCurrentUserEmail().toLowerCase();
  
  // 1. 아직 등록된 관리자가 없는 최초 설정 상태이면 설정 허용
  if (adminEmails.length === 0 && !configuredPin) {
    return { authorized: true, isInitialSetup: true, email: currentEmail };
  }
  
  // 2. 구글 이메일 일치 검증
  if (currentEmail && adminEmails.indexOf(currentEmail) !== -1) {
    return { authorized: true, by: 'email', email: currentEmail };
  }
  
  // 3. 비상 PIN 번호 일치 검증
  if (configuredPin && clientPin && String(clientPin).trim() === configuredPin.trim()) {
    return { authorized: true, by: 'pin', email: currentEmail };
  }
  
  return { authorized: false, email: currentEmail, hasAdmins: true };
}

// 클라이언트 RPC: 보안 설정 저장
function saveAdminSecurityConfig(emails, pin) {
  var currentAuth = checkAdminAuth(pin);
  if (!currentAuth.authorized && !currentAuth.isInitialSetup) {
    throw new Error('보안 설정을 변경할 관리자 권한이 없습니다.');
  }
  
  var cleanEmails = String(emails || '').split(',').map(function(s) { return s.trim().toLowerCase(); }).filter(Boolean).join(',');
  var cleanPin = String(pin || '').trim();
  
  PropertiesService.getScriptProperties().setProperties({
    'ADMIN_EMAILS': cleanEmails,
    'ADMIN_PIN': cleanPin
  });
  
  return { ok: true, adminCount: cleanEmails ? cleanEmails.split(',').length : 0, hasPin: !!cleanPin };
}

// 클라이언트 RPC: 현재 보안 상태 및 사용자 이메일 확인
function getAdminSecurityStatus(pin) {
  var auth = checkAdminAuth(pin);
  var adminEmails = getAdminEmails();
  return {
    isAuthorized: auth.authorized,
    isInitialSetup: !!auth.isInitialSetup,
    currentUserEmail: getCurrentUserEmail(),
    adminEmails: (auth.authorized || auth.isInitialSetup) ? adminEmails.join(', ') : '',
    hasPin: !!getAdminPin()
  };
}

// 클라이언트 RPC: PIN 번호로 관리자 검증
function verifyAdminPin(pin) {
  return checkAdminAuth(pin);
}

// ===== 진입점 =====
function doGet(e) {
  var view = (e && e.parameter && e.parameter.view) ? e.parameter.view.toLowerCase() : 'portal';
  var pin = (e && e.parameter && e.parameter.pin) ? e.parameter.pin : '';
  
  if (view === 'admin') {
    var auth = checkAdminAuth(pin);
    var tmpl = HtmlService.createTemplateFromFile('Admin');
    tmpl.authStatus = JSON.stringify(auth);
    return tmpl.evaluate()
      .setTitle('서울사대부중 통합 업무 관리자 콘솔')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('2027 서울사대부중 베이스캠프')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ===== 깃허브 =====
function githubHeaders() {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  const h = {
    'Accept': 'application/vnd.github+json',
    'User-Agent': 'meal-to-notion-dashboard'
  };
  if (token) h['Authorization'] = `Bearer ${token}`;
  return h;
}

// 파일 조회 (없으면 null)
function githubGetFile(path) {
  const res = UrlFetchApp.fetch(
    `https://api.github.com/repos/${GH_OWNER_REPO}/contents/${path}?ref=${GH_BRANCH}`,
    { headers: githubHeaders(), muteHttpExceptions: true }
  );
  if (res.getResponseCode() === 200) return JSON.parse(res.getContentText());
  return null;
}

// 파일 생성/갱신 (base64 문자열을 그대로 올린다)
function githubPutFile(path, base64Content, message) {
  const existing = githubGetFile(path);
  const payload = {
    message: message,
    content: base64Content,
    branch: GH_BRANCH
  };
  if (existing && existing.sha) payload.sha = existing.sha;

  const res = UrlFetchApp.fetch(
    `https://api.github.com/repos/${GH_OWNER_REPO}/contents/${path}`,
    {
      method: 'put',
      headers: githubHeaders(),
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );
  const code = res.getResponseCode();
  if (code === 200 || code === 201) {
    return { ok: true, commit: JSON.parse(res.getContentText()).commit.sha.substring(0, 7) };
  }
  return { ok: false, error: `깃허브 ${code}: ${res.getContentText().substring(0, 200)}` };
}

// ===== 노션 선택속성 =====
function notionGetSeedOptions(config) {
  const res = UrlFetchApp.fetch(
    `https://api.notion.com/v1/databases/${config.NOTION_DB_ID}`,
    {
      headers: {
        'Authorization': `Bearer ${config.NOTION_TOKEN}`,
        'Notion-Version': '2022-06-28'
      },
      muteHttpExceptions: true
    }
  );
  if (res.getResponseCode() !== 200) return null;
  const prop = JSON.parse(res.getContentText())
    .properties[config.TIMETABLE_PROP_NAME || 'SEED'];
  return (prop && prop.select) ? prop.select.options : null;
}

// 기존 옵션은 id로 보존하고 새 이름만 덧붙인다 (기존 페이지 값이 지워지지 않도록)
function notionAddSeedOption(seedName, config) {
  const options = notionGetSeedOptions(config);
  if (!options) return { ok: false, error: '노션 SEED 속성을 읽지 못했습니다' };
  if (options.some(o => o.name === seedName)) return { ok: true, already: true };

  const next = options.map(o => ({ id: o.id }));
  next.push({ name: seedName });

  const body = { properties: {} };
  body.properties[config.TIMETABLE_PROP_NAME || 'SEED'] = { select: { options: next } };

  const res = UrlFetchApp.fetch(
    `https://api.notion.com/v1/databases/${config.NOTION_DB_ID}`,
    {
      method: 'patch',
      headers: {
        'Authorization': `Bearer ${config.NOTION_TOKEN}`,
        'Notion-Version': '2022-06-28',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    }
  );
  if (res.getResponseCode() === 200) return { ok: true };
  return { ok: false, error: `노션 ${res.getResponseCode()}: ${res.getContentText().substring(0, 200)}` };
}

// ===== 학사일정 =====
// 2027 주요 학사 일정 FINAL DB
const EVENTS_DB_ID_DEFAULT = '82d2e612-fa98-823b-9c4e-011280e365af';

// 수업이 없는 날로 보이는 일정 ('방학식'은 수업일이므로 제외)
const RE_HOLIDAY = /(재량휴업|휴업일|대체\s*휴일|대체\s*공휴일|공휴일|추석|설\s*연휴|설날|한글날|개천절|광복절|삼일절|어린이날|현충일|성탄|신정)|방학(?!식)/;
// 시간표가 평소와 달라질 만한 일정
const RE_TIMETABLE = /(단축|교시|동아리|임시\s*시간표|시간표|금요일|고사|시험|행사|축제|솔밭제|체육대회|발표회|훈련|개학식|졸업식|수학여행|체험학습)/;

function eventsDbId() {
  return PropertiesService.getScriptProperties().getProperty('EVENTS_DB_ID') || EVENTS_DB_ID_DEFAULT;
}

// 그 달에 "걸치는" 일정을 모두 가져와 날짜별로 펼친다.
// 여름방학(7/16~8/16)처럼 이전 달에 시작한 일정을 놓치지 않으려면
// 시작일 기준으로 거르면 안 된다.
function getSchoolEventsByDate(yearMonth, config) {
  const year = parseInt(yearMonth.substring(0, 4), 10);
  const month = parseInt(yearMonth.substring(4, 6), 10) - 1;
  const first = new Date(year, month, 1);
  const last = new Date(year, month + 1, 0);
  const fmt = d => Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // 넉넉히 앞쪽까지 가져온 뒤 코드에서 겹침을 판정한다 (긴 기간 일정 대비)
  const lookBack = new Date(year, month - 4, 1);

  const body = {
    page_size: 100,
    filter: {
      and: [
        { property: '날짜', date: { on_or_before: fmt(last) } },
        { property: '날짜', date: { on_or_after: fmt(lookBack) } }
      ]
    }
  };

  const byDate = {};
  let cursor = null, more = true, guard = 0;

  while (more && guard++ < 10) {
    if (cursor) body.start_cursor = cursor;
    const res = UrlFetchApp.fetch(
      `https://api.notion.com/v1/databases/${eventsDbId()}/query`,
      {
        method: 'post',
        headers: {
          'Authorization': `Bearer ${config.NOTION_TOKEN}`,
          'Notion-Version': '2022-06-28',
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(body),
        muteHttpExceptions: true
      }
    );
    if (res.getResponseCode() !== 200) {
      Logger.log(`학사일정 조회 실패: ${res.getContentText().substring(0, 200)}`);
      break;
    }
    const json = JSON.parse(res.getContentText());

    (json.results || []).forEach(p => {
      const d = p.properties['날짜'] && p.properties['날짜'].date;
      if (!d || !d.start) return;
      const title = (p.properties['일정이름'] && p.properties['일정이름'].title || [])
        .map(t => t.plain_text).join('');
      if (!title) return;

      const start = d.start.substring(0, 10);
      const end = (d.end || d.start).substring(0, 10);

      // 이 달과 겹치는 구간만 펼친다
      const from = start > fmt(first) ? start : fmt(first);
      const to = end < fmt(last) ? end : fmt(last);
      if (from > to) return;

      const item = {
        name: title,
        holiday: RE_HOLIDAY.test(title),
        timetable: RE_TIMETABLE.test(title),
        span: (start !== end) ? (start + '~' + end) : ''
      };

      let cur = new Date(from + 'T00:00:00');
      const stop = new Date(to + 'T00:00:00');
      let loop = 0;
      while (cur <= stop && loop++ < 40) {
        const key = Utilities.formatDate(cur, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        if (!byDate[key]) byDate[key] = [];
        byDate[key].push(item);
        cur = new Date(cur.getTime() + 86400000);
      }
    });

    more = json.has_more;
    cursor = json.next_cursor;
  }
  return byDate;
}

// ===== SEED 이름 정규화 =====
// 파일명이자 주소의 일부가 되므로 영문 소문자·숫자·밑줄만 허용한다.
function normalizeSeedName(raw) {
  const cleaned = String(raw || '')
    .trim()
    .toLowerCase()
    .replace(/\.png$/i, '')
    .replace(/[^a-z0-9_]+/g, '_')
    .replace(/_{2,}/g, '_')
    .replace(/^_+|_+$/g, '');
  return cleaned;
}

// ===== 대시보드 조회 =====
function api_getDashboard(yearMonth) {
  const config = getConfig();
  yearMonth = yearMonth || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMM');

  const result = {
    yearMonth: yearMonth,
    syncedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    hasGithubToken: !!PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN'),
    days: [],
    seeds: [],
    warnings: [],  // 실제 위험 — 빨갛게
    notes: []      // 알아만 두면 되는 것 — 조용하게
  };

  // 저장소 시간표 목록
  const registry = listTimetableSeeds();
  const repoSeeds = sortSeeds(registry.names);

  // 노션 선택지
  const options = notionGetSeedOptions(config);
  const optionNames = options ? options.map(o => o.name) : [];
  if (!options) result.warnings.push('노션 SEED 속성을 읽지 못했습니다. 토큰이나 DB 설정을 확인하세요.');

  result.seeds = repoSeeds.map(s => ({
    name: s,
    desc: registry.desc[s] || '',
    url: imageUrlForSeed(s),
    inNotion: optionNames.indexOf(s) !== -1
  }));

  // 이번 달 노션 페이지
  const pages = getNotionPagesMapFull(yearMonth, config.TIMETABLE_PROP_NAME || 'SEED', config);
  const byDate = {};
  pages.forEach(p => { if (p.date) byDate[p.date] = p; });
  const usedSeeds = pages.map(p => p.timetableValue).filter(v => !!v);

  // 이미지가 없는 선택지 — 실제로 쓰는 날이 있을 때만 위험하다.
  // 비어 있는 자리를 매번 빨갛게 띄우면 경고 자체를 무시하게 된다.
  optionNames.forEach(n => {
    if (repoSeeds.indexOf(n) !== -1) return;
    const usedOn = pages.filter(p => p.timetableValue === n).map(p => p.date);
    if (usedOn.length) {
      result.warnings.push(`${n} — ${usedOn.join(', ')}에 지정돼 있으나 저장소에 이미지가 없습니다. 깨진 이미지가 들어갑니다.`);
    } else {
      result.notes.push(`${n} — 이미지가 아직 없는 빈 자리입니다. 이번 달에 쓰는 날은 없습니다.`);
    }
  });
  result.seeds.forEach(s => {
    if (!s.inNotion) {
      result.warnings.push(`${s.name} — 이미지는 있으나 노션 선택지에 없습니다. 날짜에 지정할 수 없습니다.`);
    }
  });

  const weekdays = getWeekdays(yearMonth);
  const eventsByDate = getSchoolEventsByDate(yearMonth, config);
  let mealCount = 0, unassigned = 0, missingPages = 0, checkDays = 0;

  weekdays.forEach(d => {
    const page = byDate[d];
    if (!page) missingPages++;
    const seed = page ? page.timetableValue : null;
    // 페이지가 없는 날은 "미지정"이 아니다 — 대응 방법이 다르므로 따로 센다
    if (page && !seed) unassigned++;
    if (page && page.menu) mealCount++;

    const events = eventsByDate[d] || [];
    const isHoliday = events.some(e => e.holiday);
    // 휴업일이 아니면서 시간표가 달라질 만한 일정이 있는 날 = 확인이 필요한 날
    const needsCheck = !isHoliday && events.some(e => e.timetable);
    if (needsCheck) checkDays++;

    result.days.push({
      date: d,
      dow: new Date(d).getDay(),
      exists: !!page,
      seed: seed,
      hasMenu: !!(page && page.menu),
      missingImage: !!seed && repoSeeds.indexOf(seed) === -1,
      events: events,
      holiday: isHoliday,
      needsCheck: needsCheck
    });
  });

  const pageCount = weekdays.length - missingPages;
  if (missingPages > 0) {
    result.warnings.push(`평일 ${weekdays.length}일 중 ${missingPages}일은 노션 페이지가 없습니다. [1. 날짜 페이지 생성]을 실행하세요.`);
  }

  result.stats = {
    weekdays: weekdays.length,
    pages: pageCount,
    missingPages: missingPages,
    meals: mealCount,
    unassigned: unassigned,
    checkDays: checkDays,
    seedCount: repoSeeds.length,
    warningCount: result.warnings.length
  };
  return result;
}

// ===== SHOOT: 등록부터 적용까지 =====
// payload: { name, desc, base64, applyDate }
function api_registerTimetable(payload) {
  const config = getConfig();
  const steps = [];
  const push = (label, ok, detail) => steps.push({ label: label, ok: ok, detail: detail || '' });

  const name = normalizeSeedName(payload.name);
  if (!name) return { ok: false, steps: steps, error: 'SEED 이름이 비어 있습니다.' };
  if (!payload.base64) return { ok: false, steps: steps, error: '이미지가 없습니다.' };

  push('이름 확인', true, `${name}.png 로 저장합니다`);

  // 1) 이미지 업로드
  const put = githubPutFile(
    `${TIMETABLE_DIR}/${name}.png`,
    payload.base64,
    `feat: ${name} 시간표 추가 (대시보드)`
  );
  if (!put.ok) { push('깃허브 업로드', false, put.error); return { ok: false, steps: steps, error: put.error }; }
  push('깃허브 업로드', true, `커밋 ${put.commit}`);

  // 2) 설명 갱신
  const descRes = updateDescriptions(name, payload.desc || '');
  push('설명 저장', descRes.ok, descRes.ok ? 'descriptions.json 갱신' : descRes.error);

  // 3) CDN 확인 — 새 파일은 반영에 시간이 걸릴 수 있어 몇 번 재시도한다
  const url = imageUrlForSeed(name);
  let cdnOk = false;
  for (let i = 0; i < 5 && !cdnOk; i++) {
    if (i > 0) Utilities.sleep(2000);
    try {
      cdnOk = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getResponseCode() === 200;
    } catch (e) { cdnOk = false; }
  }
  push('CDN 확인', cdnOk, cdnOk ? url : '아직 반영 전입니다. 잠시 후 [3. 시간표 이미지 삽입]을 실행하세요.');

  // 4) 노션 선택지 추가
  const opt = notionAddSeedOption(name, config);
  push('노션 선택지 추가', opt.ok, opt.ok ? (opt.already ? '이미 있음' : '추가됨') : opt.error);

  // 5) 설명서 갱신
  try {
    _seedCache = null; // 방금 올린 이미지가 목록에 잡히도록
    writeManualSheet();
    push('설명서 갱신', true, '스프레드시트 사용설명서 시트');
  } catch (e) {
    push('설명서 갱신', false, String(e));
  }

  // 6) 날짜 적용 (선택)
  if (payload.applyDate) {
    const applied = api_setSeedForDate(payload.applyDate, name);
    push('날짜 적용', applied.ok, applied.ok ? `${payload.applyDate} → ${name}` : applied.error);
  }

  return { ok: true, steps: steps, name: name, url: url };
}

function updateDescriptions(name, desc) {
  const path = `${TIMETABLE_DIR}/descriptions.json`;
  let map = {};
  const existing = githubGetFile(path);
  if (existing && existing.content) {
    try {
      map = JSON.parse(Utilities.newBlob(Utilities.base64Decode(existing.content)).getDataAsString());
    } catch (e) { map = {}; }
  }
  map[name] = desc || name;

  const json = JSON.stringify(map, null, 2) + '\n';
  const b64 = Utilities.base64Encode(json, Utilities.Charset.UTF_8);
  const res = githubPutFile(path, b64, `docs: ${name} 설명 갱신 (대시보드)`);
  return res.ok ? { ok: true } : { ok: false, error: res.error };
}

// ===== 특정 날짜에 SEED 지정 + 이미지 반영 =====
function api_setSeedForDate(dateStr, seedName) {
  const config = getConfig();
  const prop = config.TIMETABLE_PROP_NAME || 'SEED';
  const yearMonth = dateStr.substring(0, 4) + dateStr.substring(5, 7);

  const pages = getNotionPagesMapFull(yearMonth, prop, config);
  const page = pages.filter(p => p.date === dateStr)[0];
  if (!page) return { ok: false, error: `${dateStr} 노션 페이지가 없습니다. [1. 날짜 페이지 생성]을 먼저 실행하세요.` };

  const props = {};
  props[prop] = seedName ? { select: { name: seedName } } : { select: null };
  if (!patchNotionPage(page.id, props, config)) {
    return { ok: false, error: '노션 속성 변경에 실패했습니다.' };
  }
  writeLog('대시보드', dateStr, true, `SEED 변경: ${seedName || '(해제)'}`);

  if (!seedName) return { ok: true };

  // 이미지 교체
  const imageUrl = imageUrlForSeed(seedName);
  const blocks = getPageBlocks(page.id, config);
  const existingImage = blocks.filter(b => b.type === 'image')[0];
  if (existingImage && existingImage.image && existingImage.image.external &&
      existingImage.image.external.url === imageUrl) {
    return { ok: true, skipped: true };
  }
  if (existingImage) {
    deleteBlock(existingImage.id, config);
    Utilities.sleep(API_DELAY);
  }
  const ok = appendImageBlock(page.id, imageUrl, config);
  writeLog('대시보드', dateStr, ok, ok ? `이미지 삽입: ${seedName}` : '이미지 삽입 실패');
  return ok ? { ok: true } : { ok: false, error: '이미지 삽입에 실패했습니다.' };
}

// ===== 월간 작업 =====
function api_runStep(step, yearMonth) {
  const config = getConfig();
  try {
    if (step === 1) {
      const map = getNotionPagesMap(yearMonth, config);
      return { ok: true, count: createMonthPages(yearMonth, config, map), label: '날짜 페이지 생성' };
    }
    if (step === 2) {
      const map = getNotionPagesMap(yearMonth, config);
      return { ok: true, count: updateMealData(yearMonth, config, map), label: '급식 메뉴 업데이트' };
    }
    if (step === 3) {
      _seedCache = null;
      return { ok: true, count: updateTimetableImages(yearMonth, config), label: '시간표 이미지 삽입' };
    }
    if (step === 9) {
      _seedCache = null;
      writeManualSheet();
      return { ok: true, count: 0, label: '설명서 갱신' };
    }
    return { ok: false, error: '알 수 없는 작업입니다.' };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

// ===== 설정 =====
function api_getSettings() {
  const p = PropertiesService.getScriptProperties().getProperties();
  const mask = v => v ? (v.substring(0, 6) + '••••••' + v.substring(v.length - 4)) : '';
  return {
    githubToken: mask(p.GITHUB_TOKEN),
    notionToken: mask(p.NOTION_TOKEN),
    notionDbId: p.NOTION_DB_ID || DEFAULT_2027_CONFIG.NOTION_DB_ID,
    eventsDbId: p.EVENTS_DB_ID || DEFAULT_2027_CONFIG.EVENTS_DB_ID,
    weeklyDbId: p.WEEKLY_DB_ID || DEFAULT_2027_CONFIG.WEEKLY_DB_ID,
    todoDbId: p.TODO_DB_ID || DEFAULT_2027_CONFIG.TODO_DB_ID,
    mainPageId: p.MAIN_PAGE_ID || DEFAULT_2027_CONFIG.MAIN_PAGE_ID,
    seedProp: p.TIMETABLE_PROP_NAME || 'SEED',
    repo: GH_OWNER_REPO
  };
}

function api_saveGithubToken(token) {
  token = String(token || '').trim();
  if (!token) return { ok: false, error: '토큰이 비어 있습니다.' };

  // 저장 전에 실제로 쓰기 권한이 있는지 확인한다
  const res = UrlFetchApp.fetch(`https://api.github.com/repos/${GH_OWNER_REPO}`, {
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/vnd.github+json',
      'User-Agent': 'meal-to-notion-dashboard'
    },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    return { ok: false, error: `저장소에 접근하지 못했습니다 (${res.getResponseCode()})` };
  }
  if (!JSON.parse(res.getContentText()).permissions.push) {
    return { ok: false, error: '이 토큰에는 쓰기(push) 권한이 없습니다.' };
  }

  PropertiesService.getScriptProperties().setProperty('GITHUB_TOKEN', token);
  return { ok: true };
}

function api_saveNotionSettings(settings) {
  if (!settings) return { ok: false, error: '설정값이 비어 있습니다.' };
  const p = PropertiesService.getScriptProperties();
  if (settings.notionToken) p.setProperty('NOTION_TOKEN', settings.notionToken.trim());
  if (settings.notionDbId) p.setProperty('NOTION_DB_ID', settings.notionDbId.trim());
  if (settings.eventsDbId) p.setProperty('EVENTS_DB_ID', settings.eventsDbId.trim());
  if (settings.weeklyDbId) p.setProperty('WEEKLY_DB_ID', settings.weeklyDbId.trim());
  if (settings.todoDbId) p.setProperty('TODO_DB_ID', settings.todoDbId.trim());
  if (settings.mainPageId) p.setProperty('MAIN_PAGE_ID', settings.mainPageId.trim());
  return { ok: true };
}

// =================================================================
// 1. 주간 교육활동 계획 (AI 공문 파서 & 노션 DB 등록)
// DB: 사대부중 주간/일간 교육활동계획 (5cd2e612-fa98-8379-a493-8164aaf73794)
// =================================================================

function api_parseNoticeText(text, targetYear) {
  if (!text || !text.trim()) return { ok: false, error: '분석할 공문 텍스트를 입력해 주세요.' };
  
  targetYear = parseInt(targetYear, 10) || new Date().getFullYear();

  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
  const items = [];
  const dateRegex = /(?:(\d{1,2})[\.\/\월\-\s]+(\d{1,2})일?)(?:\s*[\(（]([월화수목금토일])[\)）])?/;
  const deptKeywords = [
    '교무기획', '교무부', '학생인성안전', '학생인성', '인성부', '교육연구', '연구부',
    '창의체험', '창체부', '과학정보', '정보부', '진로상담', '상담부', '환경복지',
    '1학년', '2학년', '3학년', '전체', '교직원', '학부모'
  ];

  let lastFoundDate = null;
  let lastDow = '';
  let idCounter = 1;

  lines.forEach(line => {
    // 제목형 머리글/공지 안내줄 필터링
    if (/^\[.*안내.*\]$|^주간.*계획$|^공문.*전달$|^\d+\.\s*주간\s*교육/i.test(line)) return;

    const match = line.match(dateRegex);
    let dateStr = lastFoundDate;
    let dow = lastDow;

    if (match) {
      const m = parseInt(match[1], 10);
      const d = parseInt(match[2], 10);
      if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
        dateStr = `${targetYear}-${('0' + m).slice(-2)}-${('0' + d).slice(-2)}`;
        dow = match[3] || ['일','월','화','수','목','금','토'][new Date(dateStr).getDay()];
        lastFoundDate = dateStr;
        lastDow = dow;
      }
    }

    if (!dateStr) return;

    // 본문에서 날짜 부분 제거
    let cleaned = line.replace(dateRegex, '').replace(/^[-*•\d\.\)\s]+/, '').trim();
    if (!cleaned) return;

    // 부서/구분 감지
    let dept = '교무기획';
    for (let k = 0; k < deptKeywords.length; k++) {
      if (line.indexOf(deptKeywords[k]) !== -1) {
        dept = deptKeywords[k].replace(/부$/, '');
        break;
      }
    }

    // 시간 및 장소/세부내용 추출
    let details = '';
    const timeMatch = cleaned.match(/(\d{1,2}:\d{2}|\d+교시)/);
    if (timeMatch) {
      details = timeMatch[1];
      cleaned = cleaned.replace(timeMatch[1], '').trim();
    }

    const parenMatch = cleaned.match(/[\(（]([^()（）]+)[\)）]/);
    if (parenMatch) {
      const extra = parenMatch[1].trim();
      details = details ? (details + ' / ' + extra) : extra;
      cleaned = cleaned.replace(/[\(（][^()（）]+[\)）]/, '').trim();
    }

    cleaned = cleaned.replace(/^[-:,\s]+|[-:,\s]+$/g, '').trim();
    if (!cleaned && details) {
      cleaned = details;
      details = '';
    }

    if (cleaned) {
      items.push({
        id: 'parsed_' + idCounter++,
        date: dateStr,
        dow: dow,
        title: cleaned,
        category: dept,
        details: details,
        selected: true
      });
    }
  });

  return { ok: true, count: items.length, items: items };
}

function api_pushWeeklyActivities(items) {
  if (!items || !items.length) return { ok: false, error: '전송할 항목이 선택되지 않았습니다.' };
  const config = getConfig();
  const dbId = config.WEEKLY_DB_ID || DEFAULT_2027_CONFIG.WEEKLY_DB_ID;
  const schema = getNotionDbSchema(dbId, config);

  const results = [];
  let success = 0, fail = 0;

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    if (!item.title || !item.date) continue;
    Utilities.sleep(API_DELAY);

    const props = {};
    props[schema.titleProp] = { title: [{ text: { content: item.title } }] };
    props[schema.dateProp] = { date: { start: item.date } };

    if (schema.selectProps && schema.selectProps.length > 0 && item.category) {
      props[schema.selectProps[0]] = { select: { name: item.category } };
    }
    if (schema.richTextProps && schema.richTextProps.length > 0 && item.details) {
      props[schema.richTextProps[0]] = { rich_text: [{ text: { content: item.details } }] };
    }

    const payload = {
      parent: { database_id: dbId },
      properties: props
    };

    try {
      const res = UrlFetchApp.fetch('https://api.notion.com/v1/pages', {
        method: 'post',
        headers: {
          'Authorization': `Bearer ${config.NOTION_TOKEN}`,
          'Content-Type': 'application/json',
          'Notion-Version': '2022-06-28'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      if (res.getResponseCode() === 200) {
        success++;
        results.push({ title: item.title, date: item.date, ok: true });
      } else {
        fail++;
        results.push({ title: item.title, date: item.date, ok: false, error: res.getContentText().substring(0, 150) });
      }
    } catch (e) {
      fail++;
      results.push({ title: item.title, date: item.date, ok: false, error: String(e) });
    }
  }

  return { ok: true, success: success, fail: fail, results: results };
}

function api_getWeeklyActivities(yearMonth) {
  const config = getConfig();
  const dbId = config.WEEKLY_DB_ID || DEFAULT_2027_CONFIG.WEEKLY_DB_ID;
  const schema = getNotionDbSchema(dbId, config);

  yearMonth = yearMonth || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMM');
  const year = parseInt(yearMonth.substring(0, 4), 10);
  const month = parseInt(yearMonth.substring(4, 6), 10) - 1;
  const first = Utilities.formatDate(new Date(year, month, 1), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const last = Utilities.formatDate(new Date(year, month + 1, 0), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const body = {
    page_size: 100,
    sorts: [{ property: schema.dateProp, direction: 'ascending' }],
    filter: {
      and: [
        { property: schema.dateProp, date: { on_or_after: first } },
        { property: schema.dateProp, date: { on_or_before: last } }
      ]
    }
  };

  try {
    const res = UrlFetchApp.fetch(`https://api.notion.com/v1/databases/${dbId}/query`, {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${config.NOTION_TOKEN}`,
        'Content-Type': 'application/json',
        'Notion-Version': '2022-06-28'
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      return { ok: false, error: `조회 실패 (${res.getResponseCode()}): ${res.getContentText().substring(0, 150)}` };
    }

    const json = JSON.parse(res.getContentText());
    const items = (json.results || []).map(p => {
      const d = p.properties[schema.dateProp]?.date?.start || '';
      const t = (p.properties[schema.titleProp]?.title || []).map(x => x.plain_text).join('');
      let cat = '';
      if (schema.selectProps && schema.selectProps.length > 0) {
        cat = p.properties[schema.selectProps[0]]?.select?.name || '';
      }
      let det = '';
      if (schema.richTextProps && schema.richTextProps.length > 0) {
        det = (p.properties[schema.richTextProps[0]]?.rich_text || []).map(x => x.plain_text).join('');
      }
      return { id: p.id, date: d, title: t, category: cat, details: det };
    });

    return { ok: true, items: items };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

// =================================================================
// 2. 단축수업 & 긴급공지 배너 원격 제어
// Page: 2027 메인 베이스캠프 (91d2e612-fa98-824b-8469-813a222ad032)
// =================================================================

function api_updateEmergencyBanner(payload) {
  const config = getConfig();
  const pageId = config.MAIN_PAGE_ID || DEFAULT_2027_CONFIG.MAIN_PAGE_ID;

  const title = payload.title || '정상 일과 운영 중';
  const details = payload.details || '';
  const color = payload.color || 'blue_background';
  const emoji = payload.emoji || '🔔';

  try {
    const blocksRes = UrlFetchApp.fetch(`https://api.notion.com/v1/blocks/${pageId}/children?page_size=30`, {
      headers: {
        'Authorization': `Bearer ${config.NOTION_TOKEN}`,
        'Notion-Version': '2022-06-28'
      },
      muteHttpExceptions: true
    });

    if (blocksRes.getResponseCode() !== 200) {
      return { ok: false, error: `메인 페이지 블록 조회 실패 (${blocksRes.getResponseCode()})` };
    }

    const blocks = JSON.parse(blocksRes.getContentText()).results || [];
    const callout = blocks.find(b => b.type === 'callout');

    const richText = [
      { type: 'text', text: { content: `${emoji} ${title}` }, annotations: { bold: true } }
    ];
    if (details) {
      richText.push({ type: 'text', text: { content: `\n${details}` } });
    }

    const calloutPayload = {
      callout: {
        rich_text: richText,
        icon: { emoji: emoji },
        color: color
      }
    };

    if (callout) {
      const patchRes = UrlFetchApp.fetch(`https://api.notion.com/v1/blocks/${callout.id}`, {
        method: 'patch',
        headers: {
          'Authorization': `Bearer ${config.NOTION_TOKEN}`,
          'Content-Type': 'application/json',
          'Notion-Version': '2022-06-28'
        },
        payload: JSON.stringify(calloutPayload),
        muteHttpExceptions: true
      });
      if (patchRes.getResponseCode() === 200) {
        PropertiesService.getScriptProperties().setProperty('SAVED_BANNER_STATE', JSON.stringify(payload));
        return { ok: true, updated: true, blockId: callout.id };
      }
      return { ok: false, error: `콜아웃 갱신 실패 (${patchRes.getResponseCode()})` };
    } else {
      const appendRes = UrlFetchApp.fetch(`https://api.notion.com/v1/blocks/${pageId}/children`, {
        method: 'patch',
        headers: {
          'Authorization': `Bearer ${config.NOTION_TOKEN}`,
          'Content-Type': 'application/json',
          'Notion-Version': '2022-06-28'
        },
        payload: JSON.stringify({ children: [{ object: 'block', type: 'callout', ...calloutPayload }] }),
        muteHttpExceptions: true
      });
      if (appendRes.getResponseCode() === 200) {
        PropertiesService.getScriptProperties().setProperty('SAVED_BANNER_STATE', JSON.stringify(payload));
        return { ok: true, created: true };
      }
      return { ok: false, error: `콜아웃 생성 실패: ${appendRes.getContentText()}` };
    }
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function api_getBannerStatus() {
  const saved = PropertiesService.getScriptProperties().getProperty('SAVED_BANNER_STATE');
  if (saved) {
    try { return { ok: true, banner: JSON.parse(saved) }; } catch (e) {}
  }
  return {
    ok: true,
    banner: {
      mode: 'normal',
      title: '정상 수업 일과 운영 (전 교시 45분)',
      details: '점심시간 12:30 ~ 13:30 · 특이사항 없음',
      color: 'blue_background',
      emoji: '🔔'
    }
  };
}

// =================================================================
// 3. 교직원 To-Do 관리
// DB: 오늘의 할일 To-Do (70c2e612-fa98-830b-a757-818914c4bd9c)
// =================================================================

function api_getTodoList() {
  const config = getConfig();
  const dbId = config.TODO_DB_ID || DEFAULT_2027_CONFIG.TODO_DB_ID;
  const schema = getNotionDbSchema(dbId, config);

  try {
    const res = UrlFetchApp.fetch(`https://api.notion.com/v1/databases/${dbId}/query`, {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${config.NOTION_TOKEN}`,
        'Content-Type': 'application/json',
        'Notion-Version': '2022-06-28'
      },
      payload: JSON.stringify({
        page_size: 40,
        sorts: [{ property: schema.dateProp, direction: 'ascending' }]
      }),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      return { ok: false, error: `To-Do 조회 실패 (${res.getResponseCode()})` };
    }

    const json = JSON.parse(res.getContentText());
    const todos = (json.results || []).map(p => {
      const title = (p.properties[schema.titleProp]?.title || []).map(x => x.plain_text).join('');
      const date = p.properties[schema.dateProp]?.date?.start || '';
      const done = !!(p.properties[schema.checkboxProp]?.checkbox);
      let dept = '';
      if (schema.selectProps && schema.selectProps.length > 0) {
        dept = p.properties[schema.selectProps[0]]?.select?.name || '';
      }
      return { id: p.id, title: title, dueDate: date, done: done, dept: dept };
    });

    return { ok: true, todos: todos };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function api_addTodo(payload) {
  if (!payload || !payload.title) return { ok: false, error: '할 일 제목을 입력해 주세요.' };
  const config = getConfig();
  const dbId = config.TODO_DB_ID || DEFAULT_2027_CONFIG.TODO_DB_ID;
  const schema = getNotionDbSchema(dbId, config);

  const props = {};
  props[schema.titleProp] = { title: [{ text: { content: payload.title } }] };
  if (payload.dueDate) {
    props[schema.dateProp] = { date: { start: payload.dueDate } };
  }
  if (schema.checkboxProp) {
    props[schema.checkboxProp] = { checkbox: false };
  }
  if (schema.selectProps && schema.selectProps.length > 0 && payload.dept) {
    props[schema.selectProps[0]] = { select: { name: payload.dept } };
  }

  try {
    const res = UrlFetchApp.fetch('https://api.notion.com/v1/pages', {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${config.NOTION_TOKEN}`,
        'Content-Type': 'application/json',
        'Notion-Version': '2022-06-28'
      },
      payload: JSON.stringify({ parent: { database_id: dbId }, properties: props }),
      muteHttpExceptions: true
    });
    if (res.getResponseCode() === 200) return { ok: true, id: JSON.parse(res.getContentText()).id };
    return { ok: false, error: `To-Do 등록 실패: ${res.getContentText().substring(0, 150)}` };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function api_toggleTodo(pageId, completed) {
  const config = getConfig();
  const dbId = config.TODO_DB_ID || DEFAULT_2027_CONFIG.TODO_DB_ID;
  const schema = getNotionDbSchema(dbId, config);

  const props = {};
  props[schema.checkboxProp] = { checkbox: !!completed };

  try {
    const res = UrlFetchApp.fetch(`https://api.notion.com/v1/pages/${pageId}`, {
      method: 'patch',
      headers: {
        'Authorization': `Bearer ${config.NOTION_TOKEN}`,
        'Content-Type': 'application/json',
        'Notion-Version': '2022-06-28'
      },
      payload: JSON.stringify({ properties: props }),
      muteHttpExceptions: true
    });
    return { ok: res.getResponseCode() === 200 };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

// =================================================================
// 4. 학사일정 마스터 (조회 & 추가 & 삭제)
// DB: 주요 학사 일정 FINAL (82d2e612-fa98-823b-9c4e-011280e365af)
// =================================================================

function api_getSchoolEventsList(yearMonth) {
  const config = getConfig();
  const dbId = eventsDbId();
  const schema = getNotionDbSchema(dbId, config);

  yearMonth = yearMonth || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMM');
  const year = parseInt(yearMonth.substring(0, 4), 10);
  const month = parseInt(yearMonth.substring(4, 6), 10) - 1;
  const first = Utilities.formatDate(new Date(year, month - 1, 1), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const last = Utilities.formatDate(new Date(year, month + 2, 0), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  try {
    const res = UrlFetchApp.fetch(`https://api.notion.com/v1/databases/${dbId}/query`, {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${config.NOTION_TOKEN}`,
        'Content-Type': 'application/json',
        'Notion-Version': '2022-06-28'
      },
      payload: JSON.stringify({
        page_size: 100,
        sorts: [{ property: schema.dateProp, direction: 'ascending' }],
        filter: {
          and: [
            { property: schema.dateProp, date: { on_or_after: first } },
            { property: schema.dateProp, date: { on_or_before: last } }
          ]
        }
      }),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) return { ok: false, error: '학사일정 조회 실패' };

    const json = JSON.parse(res.getContentText());
    const events = (json.results || []).map(p => {
      const title = (p.properties[schema.titleProp]?.title || []).map(x => x.plain_text).join('');
      const dateObj = p.properties[schema.dateProp]?.date || {};
      return {
        id: p.id,
        name: title,
        start: dateObj.start || '',
        end: dateObj.end || '',
        holiday: RE_HOLIDAY.test(title),
        timetable: RE_TIMETABLE.test(title)
      };
    });

    return { ok: true, events: events };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function api_addSchoolEvent(payload) {
  if (!payload || !payload.name || !payload.startDate) {
    return { ok: false, error: '일정 이름과 시작일을 입력해 주세요.' };
  }
  const config = getConfig();
  const dbId = eventsDbId();
  const schema = getNotionDbSchema(dbId, config);

  const datePayload = { start: payload.startDate };
  if (payload.endDate && payload.endDate !== payload.startDate) {
    datePayload.end = payload.endDate;
  }

  const props = {};
  props[schema.titleProp] = { title: [{ text: { content: payload.name } }] };
  props[schema.dateProp] = { date: datePayload };

  try {
    const res = UrlFetchApp.fetch('https://api.notion.com/v1/pages', {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${config.NOTION_TOKEN}`,
        'Content-Type': 'application/json',
        'Notion-Version': '2022-06-28'
      },
      payload: JSON.stringify({ parent: { database_id: dbId }, properties: props }),
      muteHttpExceptions: true
    });
    return { ok: res.getResponseCode() === 200 };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function api_deleteSchoolEvent(pageId) {
  const config = getConfig();
  try {
    const res = UrlFetchApp.fetch(`https://api.notion.com/v1/pages/${pageId}`, {
      method: 'patch',
      headers: {
        'Authorization': `Bearer ${config.NOTION_TOKEN}`,
        'Content-Type': 'application/json',
        'Notion-Version': '2022-06-28'
      },
      payload: JSON.stringify({ archived: true }),
      muteHttpExceptions: true
    });
    return { ok: res.getResponseCode() === 200 };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

// =================================================================
// 5. 2027 링크 마스터 관리 (구글 시트 6종 원스톱)
// =================================================================

const DEFAULT_MASTER_LINKS = [
  { id: 'afterschool_all', name: '방과후학교 출석부 (통합)', url: 'https://docs.google.com/spreadsheets/d/1XlR_nOwt-H78Fz1...', dept: '방과후학교부' },
  { id: 'afterschool_tch', name: '방과후학교 학생출석부 (강사용)', url: 'https://docs.google.com/spreadsheets/d/1B0qXp7z8uJmG9_K...', dept: '방과후학교부' },
  { id: 'emergency_contacts', name: '교직원 비상연락망 및 좌석표', url: 'https://docs.google.com/spreadsheets/d/1HkR4_5g9aL-8v7m...', dept: '교무기획부' },
  { id: 'timetable_sheet', name: '2027 1학기 시간표 (최종)', url: 'https://docs.google.com/spreadsheets/d/1Z8jM9x_k4L3w7n9...', dept: '교무기획부' },
  { id: 'club_attendance', name: '자율동아리 출석부', url: 'https://docs.google.com/spreadsheets/d/1W5qN4mP8sK_7lR3...', dept: '창의체험부' },
  { id: 'device_repair', name: '정보화기기/방송 수리요청', url: 'https://docs.google.com/spreadsheets/d/1QX5p3Yg0FXiVVir...', dept: '과학정보부' }
];

function api_getMasterLinks() {
  const saved = PropertiesService.getScriptProperties().getProperty('MASTER_LINKS_2027');
  if (saved) {
    try { return { ok: true, links: JSON.parse(saved) }; } catch (e) {}
  }
  return { ok: true, links: DEFAULT_MASTER_LINKS };
}

function api_saveMasterLinks(links) {
  if (!links) return { ok: false, error: '링크 데이터가 없습니다.' };
  PropertiesService.getScriptProperties().setProperty('MASTER_LINKS_2027', JSON.stringify(links));
  return { ok: true };
}

// =================================================================
// 6. 컴시간알리미 실시간 연동 API (2027 베이스캠프 자체 엔진)
// =================================================================

function api_getComciganTimetable(grade, classNum) {
  try {
    var g = Number(grade) || 1;
    var c = Number(classNum) || 1;
    return getComciganClassTimetable(g, c);
  } catch (err) {
    Logger.log('api_getComciganTimetable error: ' + err.message);
    return { ok: false, error: err.message };
  }
}

function api_getComciganTeacherTimetable(teacherName) {
  try {
    if (!teacherName) return { ok: false, error: '교사 성명을 입력하세요.' };
    return getComciganTeacherTimetable(teacherName);
  } catch (err) {
    Logger.log('api_getComciganTeacherTimetable error: ' + err.message);
    return { ok: false, error: err.message };
  }
}

function api_getComciganMeta() {
  try {
    return getComciganMeta();
  } catch (err) {
    Logger.log('api_getComciganMeta error: ' + err.message);
    return { ok: false, error: err.message };
  }
}

function api_refreshComciganData() {
  try {
    getComciganRawData(true);
    return { ok: true, message: '컴시간알리미 최신 시간표를 새로고침했습니다.' };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

