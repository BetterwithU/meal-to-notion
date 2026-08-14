/**
 * 학사일정 관리 대시보드 — 웹앱 서버측
 *
 * 배포: clasp deploy (appsscript.json의 webapp 설정으로 "본인만" 접근)
 * 토큰은 PropertiesService에만 두고 브라우저로 내려보내지 않는다.
 */

// ===== 진입점 =====
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('학사일정 관리 대시보드')
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
    warnings: []
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

  // 노션에는 있는데 이미지가 없는 값 = 깨진 이미지가 들어갈 위험
  optionNames.forEach(n => {
    if (repoSeeds.indexOf(n) === -1) {
      result.warnings.push(`${n} — 노션 선택지에는 있으나 저장소에 이미지가 없습니다. 이 값을 쓰면 깨진 이미지가 들어갑니다.`);
    }
  });
  result.seeds.forEach(s => {
    if (!s.inNotion) {
      result.warnings.push(`${s.name} — 이미지는 있으나 노션 선택지에 없습니다. 날짜에 지정할 수 없습니다.`);
    }
  });

  // 이번 달 노션 페이지
  const pages = getNotionPagesMapFull(yearMonth, config.TIMETABLE_PROP_NAME || 'SEED', config);
  const byDate = {};
  pages.forEach(p => { if (p.date) byDate[p.date] = p; });

  const weekdays = getWeekdays(yearMonth);
  let mealCount = 0, unassigned = 0;

  // 급식 입력 여부는 시트에서 확인 (노션 재조회보다 빠름)
  const menuByDate = readMenusFromSheet(yearMonth, config);

  weekdays.forEach(d => {
    const page = byDate[d];
    const seed = page ? page.timetableValue : null;
    if (!seed) unassigned++;
    if (menuByDate[d]) mealCount++;
    result.days.push({
      date: d,
      dow: new Date(d).getDay(),
      exists: !!page,
      seed: seed,
      missingImage: !!seed && repoSeeds.indexOf(seed) === -1
    });
  });

  result.stats = {
    weekdays: weekdays.length,
    pages: weekdays.filter(d => byDate[d]).length,
    meals: mealCount,
    unassigned: unassigned,
    seedCount: repoSeeds.length,
    warningCount: result.warnings.length
  };
  return result;
}

function readMenusFromSheet(yearMonth, config) {
  const map = {};
  try {
    const sheet = SpreadsheetApp.openById(config.SPREADSHEET_ID).getSheetByName(yearMonth);
    if (!sheet) return map;
    sheet.getDataRange().getValues().slice(1).forEach(row => {
      const d = row[0];
      const ds = (d instanceof Date)
        ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(d).substring(0, 10);
      if (ds.length === 10 && row[1]) map[ds] = String(row[1]);
    });
  } catch (e) {
    Logger.log(`readMenusFromSheet 오류: ${e}`);
  }
  return map;
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
    notionDbId: p.NOTION_DB_ID || '',
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
