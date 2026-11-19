/**
 * 서울대학교사범대학부설중학교 베이스캠프 - 컴시간알리미 자체 연동 엔진
 *
 * 외부 라이브러리/프록시 서버 없이 Google Apps Script 내부에서 100% 자체 구동
 * - 컴시간 최신 보안 체계(hour-obfuscated-v2, AES-256-CBC, X-Hour-Token) 자가 해독
 * - 일일 변동 시간표(보강·대강) 자동 감지
 * - 서버 부하 및 차단 방지를 위한 3중 캐시(CacheService + ScriptProperties)
 */

var COMCIGAN_CONFIG = {
  HOST: 'http://comci.net:4082',
  SCHOOL_CODE: '20449',          // 서울대학교사범대학부설중학교 고유 코드
  PARAM_PREFIX: '33086',         // 시간표 쿼리 접두어
  CACHE_KEY_RAW: 'COMCIGAN_RAW_DATA_V1',
  CACHE_EXPIRATION_SEC: 900,     // 15분 캐시 (컴시간 차단 방지)
  FALLBACK_PROP_KEY: 'COMCIGAN_LAST_SUCCESS_TIMETABLE'
};

/**
 * 컴시간알리미 난독화 자바스크립트 언패킹 (내부 헬퍼)
 */
function _comciganUnpackLayer(code, regex) {
  var match = code.match(regex);
  if (!match) return null;
  var a = JSON.parse(match[1]);
  var k = parseInt(match[2], 10);
  var s = '';
  for (var i = 0; i < a.length; i++) {
    if (a[i][0]) {
      s += a[i][1];
      continue;
    }
    var b = a[i][1];
    for (var j = 0; j < b.length; j += 4) {
      s += String.fromCharCode(parseInt(b.substr(j, 4), 16) ^ ((k + (j / 4) * 13) & 255));
    }
  }
  return s;
}

/**
 * 컴시간 교원 성명 및 과목명 XOR 복호화
 */
function _comciganDecodeQnmsb(ciphertext) {
  if (!ciphertext || typeof ciphertext !== 'string') return '';
  var key = 'xr6Sb[(LD:f6';
  try {
    var rawBytes = Utilities.base64Decode(ciphertext);
    var decodedBytes = [];
    for (var i = 0; i < rawBytes.length; i++) {
      decodedBytes.push(rawBytes[i] ^ key.charCodeAt(i % key.length));
    }
    return Utilities.newBlob(decodedBytes).getDataAsString('EUC-KR');
  } catch (e) {
    return String(ciphertext);
  }
}

/**
 * AES-256-CBC 시간표 데이터 패킷(enc5) 복호화
 */
function _comciganDecryptEnc5(data) {
  data = data.replace(/\x00+$/, '');
  if (data.indexOf('enc1:') !== 0 && data.indexOf('enc5:') !== 0) return data;
  if (typeof CryptoJS === 'undefined') throw new Error('CryptoJS 라이브러리가 로드되지 않았습니다.');

  var parts = data.split(':');
  if (parts.length !== 4) throw new Error('시간표 암호문 규격이 올바르지 않습니다.');

  var useBase64 = (parts[0] === 'enc5');
  var encoding = useBase64 ? CryptoJS.enc.Base64 : CryptoJS.enc.Hex;
  var binary = [];
  for (var i = 1; i <= 3; i++) {
    binary[i] = encoding.parse(parts[i]);
  }

  // 컴시간 AES 256비트 대칭키 조합
  var keyHex = (function(a) { return a[2] + a[1] + a[3] + a[0]; })([
    '9a2d929b9200fe6e',
    '2d65294192720df6',
    '87b28f89c294f669',
    '73bacdb0cda0951b'
  ]);
  var key = CryptoJS.enc.Hex.parse(keyHex);

  var decrypted = CryptoJS.AES.decrypt({ ciphertext: binary[2] }, key, {
    iv: binary[1],
    mode: CryptoJS.mode.CBC,
    padding: CryptoJS.pad.Pkcs7
  });

  return decrypted.toString(CryptoJS.enc.Utf8).replace(/\x00+$/, '');
}

/**
 * 과목/교사 인덱스 분할 헬퍼
 */
function _comciganSplitData(dat) {
  var m = String(dat).match(/\d{1,3}(?=(\d{3})+(?!\d))|\d{1,3}$/g);
  if (!m) return [0, 0];
  return m.reverse().map(Number);
}

/**
 * 정규 시간표와 일일 시간표 비교를 통한 보강/대강 판별
 */
function _comciganCheckSubstitution(daily, orig, changeAlarm) {
  var th = 0, sb = 0, isChanged = false;
  if (changeAlarm == 1) {
    if (String(daily).startsWith('>')) {
      var numPart = daily.substring(1);
      var ma = _comciganSplitData(numPart);
      th = ma[0];
      sb = ma[1] + Number(ma[2] || 0) * 1000;
      isChanged = true;
    } else {
      var ma = _comciganSplitData(daily);
      th = ma[0];
      sb = ma[1] + Number(ma[2] || 0) * 1000;
      isChanged = false;
    }
  } else {
    var ma = _comciganSplitData(daily);
    th = ma[0];
    sb = ma[1] + Number(ma[2] || 0) * 1000;
    isChanged = (orig !== daily && daily !== 0);
  }
  return { th: th, sb: sb, isChanged: isChanged };
}

/**
 * 컴시간 서버로부터 최신 원시 시간표 수신 (3중 캐싱 적용)
 */
function getComciganRawData(forceRefresh) {
  var cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    var cached = cache.get(COMCIGAN_CONFIG.CACHE_KEY_RAW);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {
        // 캐시 파싱 실패 시 원본 조회
      }
    }
  }

  try {
    // 1. 메인 뷰어 접속 (동적 토큰 및 라우트 추출)
    var stResponse = UrlFetchApp.fetch(COMCIGAN_CONFIG.HOST + '/st', {
      muteHttpExceptions: true
    });
    if (stResponse.getResponseCode() !== 200) {
      throw new Error('컴시간 서버 접속 실패 (HTTP ' + stResponse.getResponseCode() + ')');
    }

    var html = stResponse.getBlob().getDataAsString('EUC-KR');
    var tokenMatch = html.match(/__HourRequestToken\s*=\s*['"]([^'"]+)['"]/);
    var token = tokenMatch ? tokenMatch[1] : '';

    var layer1 = _comciganUnpackLayer(html, /hour-obfuscated-v2[\s\S]*?\(function\(a,k\)\{[\s\S]*?\}\)\((\[\[[\s\S]*?\]\]),\s*(\d+)\);/);
    var layer2 = layer1 ? _comciganUnpackLayer(layer1, /hour-ajax-obfuscated-v1[\s\S]*?\(function\(a,k\)\{[\s\S]*?\}\)\((\[\[[\s\S]*?\]\]),\s*(\d+)\);/) : null;
    var routeMatch = layer2 ? layer2.match(/url:'\.\/(\d+)\?'/) : null;
    var route = routeMatch ? routeMatch[1] : '362339'; // 안전한 기본 라우트

    // 2. 서울사대부중 시간표 데이터 요청
    var queryStr = COMCIGAN_CONFIG.PARAM_PREFIX + '_' + COMCIGAN_CONFIG.SCHOOL_CODE + '_0_1_enc5';
    var param = Utilities.base64Encode(queryStr);
    var url = COMCIGAN_CONFIG.HOST + '/' + route + '?' + param;

    var ttResponse = UrlFetchApp.fetch(url, {
      headers: {
        'X-Hour-Token': token,
        'Referer': COMCIGAN_CONFIG.HOST + '/st',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) SNUMS-Basecamp'
      },
      muteHttpExceptions: true
    });

    if (ttResponse.getResponseCode() !== 200) {
      throw new Error('컴시간 시간표 응답 오류 (HTTP ' + ttResponse.getResponseCode() + ')');
    }

    var ttRaw = ttResponse.getBlob().getDataAsString('EUC-KR');
    var decryptedJson = _comciganDecryptEnc5(ttRaw);
    var parsedData = JSON.parse(decryptedJson);

    // 3. 캐시에 저장 (15분 유지) 및 영구 비상 백업 저장
    var jsonStr = JSON.stringify(parsedData);
    cache.put(COMCIGAN_CONFIG.CACHE_KEY_RAW, jsonStr, COMCIGAN_CONFIG.CACHE_EXPIRATION_SEC);
    
    // 비상용 백업 프로퍼티 갱신
    PropertiesService.getScriptProperties().setProperty(
      COMCIGAN_CONFIG.FALLBACK_PROP_KEY,
      jsonStr
    );

    return parsedData;
  } catch (err) {
    Logger.log('컴시간 실시간 조회 실패, 백업 데이터 호출: ' + err.message);
    var fallback = PropertiesService.getScriptProperties().getProperty(COMCIGAN_CONFIG.FALLBACK_PROP_KEY);
    if (fallback) {
      return JSON.parse(fallback);
    }
    throw new Error('시간표 데이터를 불러올 수 없습니다: ' + err.message);
  }
}

/**
 * 특정 학급의 주간/당일 시간표 조회 (웹 프론트엔드 연동용)
 * @param {number} grade 학년 (1~3)
 * @param {number} classNum 반 (1~6)
 * @returns {Object} 학급 시간표 객체
 */
function getComciganClassTimetable(grade, classNum) {
  var rawData = getComciganRawData();
  var weekStart = _comciganDecodeQnmsb(rawData['자료112']);
  var teachers = (rawData['자료169'] || []).map(function(t) { return _comciganDecodeQnmsb(t); });
  var subjects = (rawData['자료116'] || []).map(function(s) { return _comciganDecodeQnmsb(s); });
  var periodLabels = rawData['자료172'] || ['1교시', '2교시', '3교시', '4교시', '5교시', '6교시', '7교시'];

  var schedule = {};
  for (var day = 1; day <= 5; day++) {
    schedule[day] = [];
    for (var p = 1; p <= 7; p++) {
      var orig = (rawData['자료120'] && rawData['자료120'][grade] && rawData['자료120'][grade][classNum] && rawData['자료120'][grade][classNum][day]) ? (rawData['자료120'][grade][classNum][day][p] || 0) : 0;
      var daily = (rawData['자료107'] && rawData['자료107'][grade] && rawData['자료107'][grade][classNum] && rawData['자료107'][grade][classNum][day]) ? (rawData['자료107'][grade][classNum][day][p] || 0) : 0;
      var subInfo = _comciganCheckSubstitution(daily, orig, rawData['변경알림']);

      var room = '';
      if (rawData['강의실'] === 1 && rawData['자료114'] && rawData['자료114'][grade] && rawData['자료114'][grade][classNum] && rawData['자료114'][grade][classNum][day]) {
        var roomRaw = rawData['자료114'][grade][classNum][day][p];
        if (roomRaw && roomRaw.indexOf('_') > 0) {
          room = roomRaw.split('_')[1] || '';
        }
      }

      var origMa = _comciganSplitData(orig);
      var origSubject = subInfo.isChanged ? (subjects[origMa[1]] || '') : '';

      schedule[day].push({
        period: p,
        periodLabel: periodLabels[p - 1] || (p + '교시'),
        subject: subjects[subInfo.sb] || ((orig === 0 && daily === 0) ? '' : '자습'),
        teacher: teachers[subInfo.th] || '',
        room: room,
        isChanged: subInfo.isChanged,
        originalSubject: origSubject
      });
    }
  }

  // 오늘 요일 계산 (일=0, 월=1, ... 금=5, 토=6)
  var today = new Date();
  var dayOfWeek = today.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) dayOfWeek = 1; // 주말은 월요일 기본값

  return {
    ok: true,
    schoolName: '서울대학교사범대학부설중학교',
    weekStart: weekStart,
    grade: grade,
    classNum: classNum,
    dayOfWeek: dayOfWeek,
    todayPeriods: schedule[dayOfWeek] || [],
    weeklySchedule: schedule
  };
}

/**
 * 특정 교사의 주간 수업 시간표 조회 (2027 개인화 대시보드 연동용)
 * @param {string} teacherName 교사 성명 (예: '이현*', '이현' 등)
 * @returns {Object} 교원 개인 수업 시간표 객체
 */
function getComciganTeacherTimetable(teacherName) {
  var rawData = getComciganRawData();
  var weekStart = _comciganDecodeQnmsb(rawData['자료112']);
  var teachers = (rawData['자료169'] || []).map(function(t) { return _comciganDecodeQnmsb(t); });
  var subjects = (rawData['자료116'] || []).map(function(s) { return _comciganDecodeQnmsb(s); });
  var periodLabels = rawData['자료172'] || ['1교시', '2교시', '3교시', '4교시', '5교시', '6교시', '7교시'];

  var schedule = { 1: [], 2: [], 3: [], 4: [], 5: [] };
  var cleanTarget = teacherName.replace(/\*/g, '').trim();

  var classCounts = rawData['학급수'] || [0, 8, 8, 8];
  for (var grade = 1; grade <= 3; grade++) {
    var maxClass = Math.min(8, (classCounts[grade] || 8));
    for (var cls = 1; cls <= maxClass; cls++) {
      for (var day = 1; day <= 5; day++) {
        for (var p = 1; p <= 7; p++) {
          var orig = (rawData['자료120'] && rawData['자료120'][grade] && rawData['자료120'][grade][cls] && rawData['자료120'][grade][cls][day]) ? (rawData['자료120'][grade][cls][day][p] || 0) : 0;
          var daily = (rawData['자료107'] && rawData['자료107'][grade] && rawData['자료107'][grade][cls] && rawData['자료107'][grade][cls][day]) ? (rawData['자료107'][grade][cls][day][p] || 0) : 0;
          var subInfo = _comciganCheckSubstitution(daily, orig, rawData['변경알림']);
          var tName = teachers[subInfo.th] || '';

          if (tName && (tName.indexOf(cleanTarget) !== -1 || cleanTarget.indexOf(tName.replace(/\*/g, '')) !== -1)) {
            schedule[day].push({
              period: p,
              periodLabel: periodLabels[p - 1] || (p + '교시'),
              grade: grade,
              classNum: cls,
              classLabel: grade + '-' + cls,
              subject: subjects[subInfo.sb] || '',
              isChanged: subInfo.isChanged
            });
          }
        }
      }
    }
  }

  // 교시 순서 정렬
  for (var d = 1; d <= 5; d++) {
    schedule[d].sort(function(a, b) { return a.period - b.period; });
  }

  var today = new Date();
  var dayOfWeek = today.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) dayOfWeek = 1;

  return {
    ok: true,
    schoolName: '서울대학교사범대학부설중학교',
    weekStart: weekStart,
    teacherName: teacherName,
    dayOfWeek: dayOfWeek,
    todayPeriods: schedule[dayOfWeek] || [],
    weeklySchedule: schedule
  };
}

/**
 * 컴시간 교사 59명, 학급 24개, 담임 매핑 메타데이터 반환
 */
function getComciganMeta() {
  var rawData = getComciganRawData();
  var rawTeachers = rawData['자료169'] || [];
  var teachers = [];
  for (var i = 1; i < rawTeachers.length; i++) {
    var name = _comciganDecodeQnmsb(rawTeachers[i]);
    if (name) teachers.push({ id: i, name: name });
  }

  var classes = [];
  for (var g = 1; g <= 3; g++) {
    for (var c = 1; c <= 8; c++) {
      classes.push({ grade: g, classNum: c, label: g + '학년 ' + c + '반' });
    }
  }

  var homeroomMap = {};
  for (var g = 1; g <= 3; g++) {
    for (var c = 1; c <= 8; c++) {
      var tIdx = (rawData['담임'] && rawData['담임'][g - 1]) ? rawData['담임'][g - 1][c - 1] : 0;
      if (tIdx && tIdx < 255 && rawTeachers[tIdx]) {
        var tName = _comciganDecodeQnmsb(rawTeachers[tIdx]);
        homeroomMap[g + '-' + c] = tName;
      }
    }
  }

  return {
    ok: true,
    teachers: teachers,
    classes: classes,
    homeroomMap: homeroomMap
  };
}

