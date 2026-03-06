/**
 * 급식 + 시간표 통합 관리 시스템
 * 1단계: 평일 날짜 페이지 일괄 생성
 * 2단계: 급식 메뉴 업데이트 (NEIS API → 기존 페이지 patch)
 * 3단계: 시간표 이미지 삽입 (선택속성 1~8 → 이미지 블록)
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('급식관리')
    .addItem('1. 날짜 페이지 생성', 'manualCreatePages')
    .addItem('2. 급식 메뉴 업데이트', 'manualUpdateMeals')
    .addItem('3. 시간표 이미지 삽입', 'manualUpdateTimetable')
    .addSeparator()
    .addItem('전체 실행 (1→2→3)', 'manualFullUpdate')
    .addToUi();
}

// ===== 설정 =====
// initializeConfig()는 Apps Script 에디터에서 최초 1회만 실행하세요.
// 실행 후 PropertiesService에 저장되므로 값은 코드에 남기지 않습니다.
function initializeConfig() {
  PropertiesService.getScriptProperties().setProperties({
    'NEIS_API_KEY': '여기에_NEIS_API_KEY',
    'NOTION_TOKEN': '여기에_NOTION_TOKEN',
    'NOTION_DB_ID': '여기에_NOTION_DB_ID',
    'ATPT_OFCDC_SC_CODE': '여기에_교육청코드',
    'SD_SCHUL_CODE': '여기에_학교코드',
    'SPREADSHEET_ID': '여기에_스프레드시트_ID',
    'TIMETABLE_PROP_NAME': 'SEED'
  });
  Logger.log('설정 저장 완료');
}

function getConfig() {
  return PropertiesService.getScriptProperties().getProperties();
}

// ===== 이미지 URL 매핑 =====
const TIMETABLE_IMAGES = {
  '1_45m_6p':      'https://cdn.jsdelivr.net/gh/getbetterwithme/meal-to-notion@main/timetable/1_45m_6p.png',
  '2_45m_7p':      'https://cdn.jsdelivr.net/gh/getbetterwithme/meal-to-notion@main/timetable/2_45m_7p.png',
  '3_45m_4p_club': 'https://cdn.jsdelivr.net/gh/getbetterwithme/meal-to-notion@main/timetable/3_45m_4p_club.png',
  '4_40m_6p':      'https://cdn.jsdelivr.net/gh/getbetterwithme/meal-to-notion@main/timetable/4_40m_6p.png',
  '5_40m_7p':      'https://cdn.jsdelivr.net/gh/getbetterwithme/meal-to-notion@main/timetable/5_40m_7p.png',
  '6_35m_6p':      'https://cdn.jsdelivr.net/gh/getbetterwithme/meal-to-notion@main/timetable/6_35m_6p.png',
  '7_35m_7p':      'https://cdn.jsdelivr.net/gh/getbetterwithme/meal-to-notion@main/timetable/7_35m_7p.png',
  '8_exam_3p':     'https://cdn.jsdelivr.net/gh/getbetterwithme/meal-to-notion@main/timetable/8_exam_3p.png',
  '9_exam_2p':     'https://cdn.jsdelivr.net/gh/getbetterwithme/meal-to-notion@main/timetable/9_exam_2p.png'
};

const API_DELAY = 150; // Notion rate limit: 3req/s, 150ms면 안전

// ===== 트리거 전용: 매일 실행, 20일~말일만 동작 =====
function scheduledNextMonthUpdate() {
  const now = new Date();
  const today = now.getDate();

  if (today < 20) {
    Logger.log(`[자동실행] ${today}일 — 20일 이전이므로 스킵`);
    return;
  }

  const nextMonthDate = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const yearMonth = Utilities.formatDate(nextMonthDate, Session.getScriptTimeZone(), "yyyyMM");

  Logger.log(`[자동실행] ${today}일, 타겟 연월: ${yearMonth}`);

  const config = getConfig();
  const pagesMap = getNotionPagesMap(yearMonth, config);

  createMonthPages(yearMonth, config, pagesMap);
  updateMealData(yearMonth, config, pagesMap);
  Logger.log(`[자동실행] ${yearMonth} 완료`);
}

// ===== 수동 실행 (UI) =====
function manualCreatePages() {
  const yearMonth = promptYearMonth();
  if (!yearMonth) return;
  const config = getConfig();
  const pagesMap = getNotionPagesMap(yearMonth, config);
  const count = createMonthPages(yearMonth, config, pagesMap);
  SpreadsheetApp.getUi().alert(`${yearMonth} 날짜 페이지 ${count}개 생성 완료`);
}

function manualUpdateMeals() {
  const yearMonth = promptYearMonth();
  if (!yearMonth) return;
  const config = getConfig();
  const pagesMap = getNotionPagesMap(yearMonth, config);
  const count = updateMealData(yearMonth, config, pagesMap);
  SpreadsheetApp.getUi().alert(`${yearMonth} 급식 메뉴 ${count}건 업데이트 완료`);
}

function manualUpdateTimetable() {
  const yearMonth = promptYearMonth();
  if (!yearMonth) return;
  const config = getConfig();
  const count = updateTimetableImages(yearMonth, config);
  SpreadsheetApp.getUi().alert(`${yearMonth} 시간표 이미지 ${count}건 삽입 완료`);
}

function manualFullUpdate() {
  const yearMonth = promptYearMonth();
  if (!yearMonth) return;
  const ui = SpreadsheetApp.getUi();
  const config = getConfig();

  // pagesMap 1번만 조회, 1단계 후 갱신
  let pagesMap = getNotionPagesMap(yearMonth, config);
  const pages = createMonthPages(yearMonth, config, pagesMap);

  // 1단계에서 새 페이지가 생겼으면 맵 갱신
  if (pages > 0) pagesMap = getNotionPagesMap(yearMonth, config);

  const meals = updateMealData(yearMonth, config, pagesMap);
  const images = updateTimetableImages(yearMonth, config);

  ui.alert(`${yearMonth} 전체 완료\n- 페이지 생성: ${pages}건\n- 급식 업데이트: ${meals}건\n- 시간표 이미지: ${images}건`);
}

function promptYearMonth() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('연월을 입력하세요 (예: 202603)');
  if (response.getSelectedButton() === ui.Button.CANCEL) return null;
  const yearMonth = response.getResponseText().trim();
  if (!/^\d{6}$/.test(yearMonth)) {
    ui.alert('YYYYMM 형식으로 입력해주세요.');
    return null;
  }
  return yearMonth;
}

// ===== 1단계: 평일 날짜 페이지 생성 + 시트 기록 =====
function createMonthPages(yearMonth, config, existingMap) {
  config = config || getConfig();
  existingMap = existingMap || getNotionPagesMap(yearMonth, config);
  const weekdays = getWeekdays(yearMonth);

  // 스프레드시트에 해당 월 시트 생성
  const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(yearMonth);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(yearMonth);
    sheet.appendRow(['날짜', '메뉴', '연월', '노션페이지ID', '시간표']);
  }

  // 시트에 이미 있는 날짜 확인
  const sheetData = sheet.getDataRange().getValues();
  const sheetDates = new Set(sheetData.slice(1).map(row => {
    const d = row[0];
    return (d instanceof Date)
      ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : String(d).substring(0, 10);
  }).filter(d => d.length === 10));

  let created = 0;

  for (const dateStr of weekdays) {
    let pageId = existingMap[dateStr];
    // 요일에 따라 시간표 기본값: 월수금(1,3,5)=6교시, 화목(2,4)=7교시
    const dayOfWeek = new Date(dateStr).getDay();
    const timetableDefault = (dayOfWeek === 1 || dayOfWeek === 3 || dayOfWeek === 5) ? '1_45m_6p' : '2_45m_7p';

    if (!pageId) {
      pageId = createNotionPage(dateStr, yearMonth, timetableDefault, config);
      if (pageId) {
        created++;
        existingMap[dateStr] = pageId;
      }
      Utilities.sleep(API_DELAY);
    } else {
      // 기존 페이지: SEED 속성을 요일 기본값으로 업데이트
      const timetableProp = config.TIMETABLE_PROP_NAME || 'SEED';
      const props = {};
      props[timetableProp] = { select: { name: timetableDefault } };
      patchNotionPage(pageId, props, config);
      Utilities.sleep(API_DELAY);
    }

    if (!sheetDates.has(dateStr) && pageId) {
      sheet.appendRow([dateStr, '', yearMonth, pageId, timetableDefault]);
    }
  }

  SpreadsheetApp.flush(); // 시트 데이터 즉시 반영
  Logger.log(`[1단계] ${yearMonth} 페이지 ${created}개 생성 (기존 ${Object.keys(existingMap).length - created}개)`);
  return created;
}

// ===== 2단계: 급식 메뉴 업데이트 =====
function updateMealData(yearMonth, config, pagesMap) {
  config = config || getConfig();
  const meals = getMonthlyMeals(yearMonth, config);
  if (!meals || meals.length === 0) {
    Logger.log(`[2단계] ${yearMonth} 급식 데이터 없음`);
    return 0;
  }

  pagesMap = pagesMap || getNotionPagesMap(yearMonth, config);

  // 시트에도 메뉴 반영
  const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(yearMonth);
  let sheetData = null;
  let sheetDateMap = {};
  if (sheet) {
    sheetData = sheet.getDataRange().getValues();
    for (let i = 1; i < sheetData.length; i++) {
      const d = sheetData[i][0];
      const dateStr = (d instanceof Date)
        ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd")
        : String(d).substring(0, 10);
      sheetDateMap[dateStr] = i + 1; // 행 번호 (1-indexed)
    }
  }

  let updated = 0;

  for (const meal of meals) {
    const pageId = pagesMap[meal.date];
    if (!pageId) {
      Logger.log(`[2단계] ${meal.date} 페이지 없음, 스킵`);
      continue;
    }

    const success = patchNotionPage(pageId, {
      "메뉴": { rich_text: [{ text: { content: meal.menu } }] }
    }, config);

    if (success) {
      updated++;
      // 시트에도 메뉴 기록
      if (sheet && sheetDateMap[meal.date]) {
        sheet.getRange(sheetDateMap[meal.date], 2).setValue(meal.menu);
      }
    }
    Utilities.sleep(API_DELAY);
  }

  Logger.log(`[2단계] ${yearMonth} 급식 ${updated}건 업데이트`);
  return updated;
}

// ===== 3단계: 시간표 이미지 삽입 =====
// getNotionPagesMap에서 시간표 속성도 함께 가져와서 개별 조회 제거
function updateTimetableImages(yearMonth, config) {
  config = config || getConfig();
  const timetableProp = config.TIMETABLE_PROP_NAME || '시간표';
  const pagesWithTimetable = getNotionPagesMapFull(yearMonth, timetableProp, config);
  let inserted = 0;

  for (const page of pagesWithTimetable) {
    if (!page.timetableValue) continue; // 시간표 미지정

    const imageUrl = TIMETABLE_IMAGES[page.timetableValue];
    if (!imageUrl) continue;

    // 기존 이미지 블록 확인
    const existingBlocks = getPageBlocks(page.id, config);
    const existingImage = existingBlocks.find(b => b.type === 'image');

    // 이미 같은 이미지면 스킵
    if (existingImage && existingImage.image?.external?.url === imageUrl) continue;

    // 기존 이미지 삭제
    if (existingImage) {
      deleteBlock(existingImage.id, config);
      Utilities.sleep(API_DELAY);
    }

    // 새 이미지 삽입
    const success = appendImageBlock(page.id, imageUrl, config);
    if (success) inserted++;
    Utilities.sleep(API_DELAY);
  }

  Logger.log(`[3단계] ${yearMonth} 시간표 이미지 ${inserted}건 삽입`);
  return inserted;
}

// ===== 헬퍼: 평일 날짜 계산 =====
function getWeekdays(yearMonth) {
  const year = parseInt(yearMonth.substring(0, 4));
  const month = parseInt(yearMonth.substring(4, 6)) - 1;
  const dates = [];
  const lastDay = new Date(year, month + 1, 0).getDate();

  for (let d = 1; d <= lastDay; d++) {
    const date = new Date(year, month, d);
    const day = date.getDay();
    if (day >= 1 && day <= 5) {
      const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
      dates.push(dateStr);
    }
  }
  return dates;
}

// ===== 헬퍼: Notion DB 페이지 맵 조회 (날짜 → pageId) =====
function getNotionPagesMap(yearMonth, config) {
  const url = `https://api.notion.com/v1/databases/${config.NOTION_DB_ID}/query`;
  const map = {};
  let hasMore = true;
  let startCursor = undefined;

  while (hasMore) {
    const body = {
      page_size: 100,
      filter: { property: "Month", select: { equals: String(yearMonth) } }
    };
    if (startCursor) body.start_cursor = startCursor;

    const options = {
      method: "post",
      headers: {
        "Authorization": `Bearer ${config.NOTION_TOKEN}`,
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    };

    try {
      const res = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
      (res.results || []).forEach(p => {
        const dateStart = p.properties["날짜"]?.date?.start;
        if (dateStart) map[dateStart] = p.id;
      });
      hasMore = res.has_more;
      startCursor = res.next_cursor;
    } catch (e) {
      Logger.log(`getNotionPagesMap 오류: ${e}`);
      break;
    }
  }
  return map;
}

// ===== 헬퍼: Notion DB 페이지 + 시간표 속성 함께 조회 (3단계용) =====
function getNotionPagesMapFull(yearMonth, timetableProp, config) {
  const url = `https://api.notion.com/v1/databases/${config.NOTION_DB_ID}/query`;
  const pages = [];
  let hasMore = true;
  let startCursor = undefined;

  while (hasMore) {
    const body = {
      page_size: 100,
      filter: { property: "Month", select: { equals: String(yearMonth) } }
    };
    if (startCursor) body.start_cursor = startCursor;

    const options = {
      method: "post",
      headers: {
        "Authorization": `Bearer ${config.NOTION_TOKEN}`,
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    };

    try {
      const res = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
      (res.results || []).forEach(p => {
        const dateStart = p.properties["날짜"]?.date?.start;
        const timetableSelect = p.properties[timetableProp]?.select;
        pages.push({
          id: p.id,
          date: dateStart,
          timetableValue: timetableSelect ? timetableSelect.name : null
        });
      });
      hasMore = res.has_more;
      startCursor = res.next_cursor;
    } catch (e) {
      Logger.log(`getNotionPagesMapFull 오류: ${e}`);
      break;
    }
  }
  return pages;
}

// ===== 헬퍼: Notion 페이지 생성 (빈 페이지) =====
function createNotionPage(dateStr, yearMonth, timetableValue, config) {
  const url = 'https://api.notion.com/v1/pages';
  const timetableProp = config.TIMETABLE_PROP_NAME || '시간표';
  const properties = {
    "이름": { title: [{ text: { content: dateStr } }] },
    "날짜": { date: { start: dateStr } },
    "Month": { select: { name: String(yearMonth) } }
  };
  if (timetableValue) {
    properties[timetableProp] = { select: { name: String(timetableValue) } };
  }
  const payload = {
    parent: { database_id: config.NOTION_DB_ID },
    properties: properties
  };
  const options = {
    method: "post",
    headers: {
      "Authorization": `Bearer ${config.NOTION_TOKEN}`,
      "Content-Type": "application/json",
      "Notion-Version": "2022-06-28"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    const res = UrlFetchApp.fetch(url, options);
    if (res.getResponseCode() === 200) return JSON.parse(res.getContentText()).id;
    Logger.log(`createNotionPage 실패 (${dateStr}): ${res.getContentText()}`);
    return null;
  } catch (e) { return null; }
}

// ===== 헬퍼: Notion 페이지 속성 업데이트 =====
function patchNotionPage(pageId, properties, config) {
  const url = `https://api.notion.com/v1/pages/${pageId}`;
  const options = {
    method: "patch",
    headers: {
      "Authorization": `Bearer ${config.NOTION_TOKEN}`,
      "Content-Type": "application/json",
      "Notion-Version": "2022-06-28"
    },
    payload: JSON.stringify({ properties: properties }),
    muteHttpExceptions: true
  };
  try {
    const res = UrlFetchApp.fetch(url, options);
    return res.getResponseCode() === 200;
  } catch (e) { return false; }
}

// ===== 헬퍼: 블록 삭제 =====
function deleteBlock(blockId, config) {
  const url = `https://api.notion.com/v1/blocks/${blockId}`;
  const options = {
    method: "delete",
    headers: {
      "Authorization": `Bearer ${config.NOTION_TOKEN}`,
      "Notion-Version": "2022-06-28"
    },
    muteHttpExceptions: true
  };
  try {
    const res = UrlFetchApp.fetch(url, options);
    return res.getResponseCode() === 200;
  } catch (e) { return false; }
}

// ===== 헬퍼: 페이지 블록 목록 조회 =====
function getPageBlocks(pageId, config) {
  const url = `https://api.notion.com/v1/blocks/${pageId}/children?page_size=100`;
  const options = {
    method: "get",
    headers: {
      "Authorization": `Bearer ${config.NOTION_TOKEN}`,
      "Notion-Version": "2022-06-28"
    },
    muteHttpExceptions: true
  };
  try {
    const res = UrlFetchApp.fetch(url, options);
    if (res.getResponseCode() === 200) return JSON.parse(res.getContentText()).results || [];
    return [];
  } catch (e) { return []; }
}

// ===== 헬퍼: 이미지 블록 추가 =====
function appendImageBlock(pageId, imageUrl, config) {
  const url = `https://api.notion.com/v1/blocks/${pageId}/children`;
  const options = {
    method: "patch",
    headers: {
      "Authorization": `Bearer ${config.NOTION_TOKEN}`,
      "Content-Type": "application/json",
      "Notion-Version": "2022-06-28"
    },
    payload: JSON.stringify({
      children: [{
        object: "block",
        type: "image",
        image: {
          type: "external",
          external: { url: imageUrl }
        }
      }]
    }),
    muteHttpExceptions: true
  };
  try {
    const res = UrlFetchApp.fetch(url, options);
    return res.getResponseCode() === 200;
  } catch (e) { return false; }
}

// ===== NEIS 급식 데이터 조회 =====
function getMonthlyMeals(yearMonth, config) {
  const url = `https://open.neis.go.kr/hub/mealServiceDietInfo?KEY=${config.NEIS_API_KEY}&Type=json&ATPT_OFCDC_SC_CODE=${config.ATPT_OFCDC_SC_CODE}&SD_SCHUL_CODE=${config.SD_SCHUL_CODE}&MLSV_FROM_YMD=${yearMonth}01&MLSV_TO_YMD=${yearMonth}31`;
  try {
    const res = JSON.parse(UrlFetchApp.fetch(url).getContentText());
    if (!res.mealServiceDietInfo) return null;

    return res.mealServiceDietInfo[1].row.map(r => {
      const menuList = r.DDISH_NM.split('<br/>');
      const cleanedItems = menuList.map(item => {
        return item
          .replace(/\([^)]+\)/g, "")
          .replace(/[@*#]/g, "")
          .replace(/\s+/g, " ")
          .trim();
      }).filter(item => item.length > 0);

      return {
        date: `${r.MLSV_YMD.substring(0,4)}-${r.MLSV_YMD.substring(4,6)}-${r.MLSV_YMD.substring(6,8)}`,
        menu: cleanedItems.join(", ")
      };
    });
  } catch (e) { return null; }
}

// ===== 사용설명서 시트 생성 =====
function writeManualSheet() {
  const config = getConfig();
  const ss = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  let sheet = ss.getSheetByName('사용설명서');
  if (sheet) sheet.clear();
  else sheet = ss.insertSheet('사용설명서', 0);

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 500);

  const hdrBg = '#1a73e8';
  const hdrFont = '#ffffff';
  const secBg = '#e8f0fe';
  const warnBg = '#fff8e1';

  let r = 1;

  // 제목
  sheet.getRange(r, 1, 1, 2).merge()
    .setValue('급식+시간표 통합 관리 시스템 사용 설명서')
    .setFontSize(16).setFontWeight('bold').setBackground('#f1f3f4')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(r, 40);
  r += 2;

  // --- 새 컴퓨터에서 이어서 작업하기 ---
  r = writeSection(sheet, r, '새 컴퓨터에서 이어서 작업하기', '#fce8e6');
  r = writeRow(sheet, r, '1단계: 저장소 가져오기', '처음: git clone https://github.com/getbetterwithme/meal-to-notion.git\n이미 있으면: git pull origin main', true);
  r = writeRow(sheet, r, '2단계: clasp 설치 및 로그인', 'npm install -g @google/clasp\nclasp login  →  브라우저에서 구글 계정 허용', true);
  r = writeRow(sheet, r, '3단계: .clasp.json 만들기', 'apps-script/ 폴더 안에 .clasp.json 파일 직접 생성 (GitHub에 안 올라감)\n내용: { "scriptId": "1EH5MzU64_AxOQhJUdROvBFZz51kNNNFrIpN_U_LPiF3xgefYzyU0mDPw", "rootDir": "." }', true);
  r = writeRow(sheet, r, '4단계: Apps Script 반영', 'cd apps-script\nclasp push\n→ 이후 코드 수정 시마다 clasp push로 반영', true);
  r++;

  // --- 메뉴 구성 ---
  r = writeSection(sheet, r, '메뉴 구성', secBg);
  r = writeTableHeader(sheet, r, ['메뉴', '기능'], hdrBg, hdrFont);
  r = writeTableRow(sheet, r, ['1. 날짜 페이지 생성', '해당 월 평일(월~금) 페이지를 노션 DB에 생성\n월수금→SEED "1_45m_6p", 화목→SEED "2_45m_7p" 자동 입력']);
  r = writeTableRow(sheet, r, ['2. 급식 메뉴 업데이트', 'NEIS API에서 급식 데이터를 가져와 기존 노션 페이지에 메뉴 업데이트']);
  r = writeTableRow(sheet, r, ['3. 시간표 이미지 삽입', '각 페이지의 SEED 속성값을 읽어 해당 이미지를 페이지에 삽입/교체']);
  r = writeTableRow(sheet, r, ['전체 실행 (1→2→3)', '위 3단계를 순서대로 한 번에 실행']);
  r++;

  // --- 기본 워크플로우 ---
  r = writeSection(sheet, r, '기본 워크플로우', secBg);
  r = writeRow(sheet, r, 'A. 매월 새 달 준비', '1. [급식관리] > [전체 실행] 클릭\n2. 연월 입력 (예: 202603)\n3. 자동으로 날짜 생성 → 급식 입력 → 시간표 이미지 삽입', true);
  r = writeRow(sheet, r, 'B. 시간표 수정 시', '1. 노션 DB에서 해당 날짜의 SEED 속성값 변경 (예: 2_45m_7p → 4_40m_6p)\n2. [급식관리] > [3. 시간표 이미지 삽입] 실행\n3. 변경된 값에 해당하는 이미지로 자동 교체됨', true);
  r++;

  // --- 자동 트리거 ---
  r = writeSection(sheet, r, '자동 트리거', secBg);
  r = writeRow(sheet, r, '실행 조건', '매월 20일~말일 사이 매일 자동 실행\n다음 달 날짜 생성 + 급식 업데이트 수행\n20일 이전에는 자동 스킵', false);
  r++;

  // --- 시간표 이미지 매핑 ---
  r = writeSection(sheet, r, '시간표 이미지 매핑 (SEED 속성)', secBg);
  r = writeTableHeader(sheet, r, ['SEED 값', '기본 요일', '수업 형태'], hdrBg, hdrFont);
  r = writeTableRow(sheet, r, ['1_45m_6p', '월, 수, 금', '45분 × 6교시']);
  r = writeTableRow(sheet, r, ['2_45m_7p', '화, 목', '45분 × 7교시']);
  r = writeTableRow(sheet, r, ['3_45m_4p_club', '수동 지정', '45분 × 4교시 + 동아리']);
  r = writeTableRow(sheet, r, ['4_40m_6p', '수동 지정', '40분 × 6교시']);
  r = writeTableRow(sheet, r, ['5_40m_7p', '수동 지정', '40분 × 7교시']);
  r = writeTableRow(sheet, r, ['6_35m_6p', '수동 지정', '35분 × 6교시']);
  r = writeTableRow(sheet, r, ['7_35m_7p', '수동 지정', '35분 × 7교시']);
  r = writeTableRow(sheet, r, ['8_exam_3p', '수동 지정', '시험 × 3교시 (이미지 준비 중)']);
  r = writeTableRow(sheet, r, ['9_exam_2p', '수동 지정', '시험 × 2교시 (이미지 준비 중)']);
  sheet.getRange(r, 1, 1, 2).merge()
    .setValue('이미지 저장소: GitHub getbetterwithme/meal-to-notion/timetable/')
    .setFontSize(9).setFontColor('#888888');
  r += 2;

  // --- 주의 사항 ---
  r = writeSection(sheet, r, '주의 사항', warnBg);
  r = writeRow(sheet, r, '필수 속성', '이름(제목), 날짜(날짜), 메뉴(텍스트), Month(선택), SEED(선택, 값: 1_45m_6p 등)', false);
  r = writeRow(sheet, r, '급식 없는 달', 'NEIS에 데이터 미등록 시 2단계 0건. 추후 재실행', false);
  r = writeRow(sheet, r, '권한 오류', '스크립트 첫 실행 시 구글 권한 승인 팝업에서 허용 필요', false);
  r++;

  // --- 문제 해결 ---
  r = writeSection(sheet, r, '문제 해결', secBg);
  r = writeTableHeader(sheet, r, ['증상', '해결'], hdrBg, hdrFont);
  r = writeTableRow(sheet, r, ['페이지 생성 실패', '노션 DB ID와 API 토큰 확인 → initializeConfig() 재실행']);
  r = writeTableRow(sheet, r, ['급식 데이터 없음', 'NEIS API 키와 학교 코드 확인']);
  r = writeTableRow(sheet, r, ['이미지 미표시', 'GitHub 저장소에 이미지 파일 확인']);
  r++;

  // --- 링크 ---
  r = writeSection(sheet, r, '관련 링크', secBg);
  r = writeRow(sheet, r, '노션 데이터베이스', 'https://www.notion.so/juneywooky/2fecd3403dc38007a150fe08df7c5e8f', false);
  r = writeRow(sheet, r, 'GitHub 저장소', 'https://github.com/getbetterwithme/meal-to-notion', false);
  r = writeRow(sheet, r, 'Apps Script', 'https://script.google.com/d/1EH5MzU64_AxOQhJUdROvBFZz51kNNNFrIpN_U_LPiF3xgefYzyU0mDPw/edit', false);

  SpreadsheetApp.flush();
  Logger.log('사용설명서 작성 완료');
}

function writeSection(sheet, r, title, bg) {
  sheet.getRange(r, 1, 1, 2).merge()
    .setValue(title)
    .setFontSize(13).setFontWeight('bold').setBackground(bg)
    .setVerticalAlignment('middle');
  sheet.setRowHeight(r, 32);
  return r + 1;
}

function writeTableHeader(sheet, r, cols, bg, fontColor) {
  cols.forEach((col, i) => {
    sheet.getRange(r, i + 1).setValue(col)
      .setFontWeight('bold').setBackground(bg).setFontColor(fontColor).setFontSize(10);
  });
  return r + 1;
}

function writeTableRow(sheet, r, cols) {
  cols.forEach((col, i) => {
    sheet.getRange(r, i + 1).setValue(col).setFontSize(10).setWrap(true);
  });
  return r + 1;
}

function writeRow(sheet, r, label, desc, isBold) {
  sheet.getRange(r, 1).setValue(label).setFontWeight(isBold ? 'bold' : 'normal').setFontSize(10);
  sheet.getRange(r, 2).setValue(desc).setFontSize(10).setWrap(true);
  return r + 1;
}
