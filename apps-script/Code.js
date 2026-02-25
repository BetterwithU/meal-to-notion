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
    .addItem('날짜+급식 실행 (1→2)', 'manualFullUpdate')
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
    'TIMETABLE_PROP_NAME': '시간표'
  });
  Logger.log('설정 저장 완료');
}

function getConfig() {
  return PropertiesService.getScriptProperties().getProperties();
}

// ===== 이미지 URL 매핑 =====
const TIMETABLE_IMAGES = {
  '1': 'https://raw.githubusercontent.com/getbetterwithme/meal-to-notion/main/timetable/1.png',
  '2': 'https://raw.githubusercontent.com/getbetterwithme/meal-to-notion/main/timetable/2.png',
  '3': 'https://raw.githubusercontent.com/getbetterwithme/meal-to-notion/main/timetable/3.png',
  '4': 'https://raw.githubusercontent.com/getbetterwithme/meal-to-notion/main/timetable/4.png',
  '5': 'https://raw.githubusercontent.com/getbetterwithme/meal-to-notion/main/timetable/5.png',
  '6': 'https://raw.githubusercontent.com/getbetterwithme/meal-to-notion/main/timetable/6.png',
  '7': 'https://raw.githubusercontent.com/getbetterwithme/meal-to-notion/main/timetable/7.png',
  '8': 'https://raw.githubusercontent.com/getbetterwithme/meal-to-notion/main/timetable/8.png'
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

  ui.alert(`${yearMonth} 완료\n- 페이지 생성: ${pages}건\n- 급식 업데이트: ${meals}건\n\n시간표 이미지는 노션에서 속성 입력 후 [3. 시간표 이미지 삽입]을 별도 실행하세요.`);
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
    if (!pageId) {
      pageId = createNotionPage(dateStr, yearMonth, config);
      if (pageId) {
        created++;
        existingMap[dateStr] = pageId;
      }
      Utilities.sleep(API_DELAY);
    }

    if (!sheetDates.has(dateStr) && pageId) {
      sheet.appendRow([dateStr, '', yearMonth, pageId, '']);
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

    // 이미 이미지 블록이 있는지 확인
    const existingBlocks = getPageBlocks(page.id, config);
    const hasImage = existingBlocks.some(b => b.type === 'image');
    if (hasImage) continue;

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
function createNotionPage(dateStr, yearMonth, config) {
  const url = 'https://api.notion.com/v1/pages';
  const payload = {
    parent: { database_id: config.NOTION_DB_ID },
    properties: {
      "이름": { title: [{ text: { content: dateStr } }] },
      "날짜": { date: { start: dateStr } },
      "Month": { select: { name: String(yearMonth) } }
    }
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
