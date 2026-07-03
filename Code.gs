/**
 * Decision (가족 중대 의사결정 플랫폼) - Google Apps Script Backend
 */

var CONFIG = {
  SPREADSHEET_ID: "", // 특정 Spreadsheet ID (비어있으면 Active Spreadsheet 또는 자동 생성)
  SHEET_NAMES: {
    ISSUES: "Issues",
    COMMENTS: "Comments",
    VOTES: "Votes",
    LOGS: "Logs",
    MEMBERS: "Members"
  }
};

// Fail-safe getDb() - NEVER returns null
function getDb() {
  var ss = null;
  var prop = PropertiesService.getScriptProperties();
  var savedId = prop.getProperty("SPREADSHEET_ID");

  // 1. CONFIG.SPREADSHEET_ID가 설정되어 있는 경우
  if (CONFIG.SPREADSHEET_ID && CONFIG.SPREADSHEET_ID.trim() !== "") {
    try { ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID.trim()); } catch (e) {}
  }

  // 2. 이미 생성되어 저장된 Script Properties ID가 있는 경우
  if (!ss && savedId) {
    try { ss = SpreadsheetApp.openById(savedId); } catch (e) {}
  }

  // 3. 바운드 스크립트로 실행되어 ActiveSpreadsheet가 있는 경우
  if (!ss) {
    try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  }

  // 4. 위 모든 시도가 실패 시 새 구글 시트 자동 생성 후 ID 저장
  if (!ss) {
    ss = SpreadsheetApp.create("Decision_DB");
    prop.setProperty("SPREADSHEET_ID", ss.getId());
  }

  ensureSheetsAndHeaders(ss);
  return ss;
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    var action = e.parameter.action;
    var result = handleApiRequest(action, e.parameter);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var template = HtmlService.createTemplateFromFile("Index");
  return template.evaluate()
    .setTitle("Decision - 의사결정 플랫폼")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  var requestData = {};
  try {
    if (e && e.postData && e.postData.contents) {
      requestData = JSON.parse(e.postData.contents);
    }
  } catch (err) {
    requestData = {};
  }

  var action = requestData.action || (e && e.parameter ? e.parameter.action : "");
  var payload = requestData.payload || requestData;

  var result = handleApiRequest(action, payload);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleApiRequest(action, payload) {
  try {
    switch (action) {
      case 'getDashboardData':
        return { success: true, data: getDashboardData() };
      case 'getIssues':
        return { success: true, data: getIssues(payload.filters || payload) };
      case 'getIssueDetail':
        return { success: true, data: getIssueDetail(payload.issueId || payload.id) };
      case 'createIssue':
        return { success: true, data: createIssue(payload) };
      case 'addComment':
        return { success: true, data: addComment(payload.issueId, payload.author, payload.text) };
      case 'castVote':
        return { success: true, data: castVote(payload.issueId, payload.voter, payload.value, payload.reason) };
      case 'updateIssueStatus':
        return { success: true, data: updateIssueStatus(payload.issueId, payload.actor, payload.toStatus, payload.holdReason, payload.reviewDate) };
      case 'deleteIssue':
        return { success: true, data: deleteIssue(payload.issueId, payload.actor) };
      case 'generateMeetingDoc':
        return { success: true, data: generateMeetingDoc(payload.issueId) };
      default:
        throw new Error("알 수 없는 API action입니다: " + action);
    }
  } catch (err) {
    return { error: true, message: err.message || String(err) };
  }
}

function ensureSheetsAndHeaders(ss) {
  var schema = {
    Issues: [
      'id', 'title', 'category', 'proposer', 'decisionOwner', 'executor', 
      'deadline', 'estimatedCost', 'summary', 'risk', 'optionA', 'optionB', 
      'optionC', 'status', 'priority', 'holdReason', 'reviewDate', 'docUrl', 
      'createdAt', 'updatedAt', 'deleted'
    ],
    Comments: ['id', 'issueId', 'author', 'text', 'createdAt'],
    Votes: ['id', 'issueId', 'voter', 'value', 'reason', 'createdAt'],
    Logs: ['id', 'issueId', 'actor', 'fromStatus', 'toStatus', 'reason', 'createdAt'],
    Members: ['id', 'name', 'role', 'canApprove', 'createdAt']
  };

  Object.keys(schema).forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(schema[sheetName]);
      sheet.getRange(1, 1, 1, schema[sheetName].length).setFontWeight("bold").setBackground("#f3f4f6");
    }
  });

  var issuesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ISSUES);
  if (issuesSheet.getLastRow() <= 1) {
    seedInitialData(ss);
  }
}

function seedInitialData(ss) {
  var now = new Date().toISOString();
  var membersSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MEMBERS);
  if (membersSheet.getLastRow() <= 1) {
    var initialMembers = [
      ['MEM_1', '이몽룡', '최종결정권자 / 대표', true, now],
      ['MEM_2', '홍길동', '기획담당', false, now],
      ['MEM_3', '성춘향', '운영담당', false, now]
    ];
    initialMembers.forEach(function(row) { membersSheet.appendRow(row); });
  }

  var issuesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ISSUES);
  var initialIssues = [
    [
      'ISSUE_101', '신규 프로젝트 추진의 건', '기획', '홍길동', '이몽룡', '성춘향',
      '7월 15일', '500만원',
      '신규 서비스 도입을 위한 시장 조사 및 사전 개발 범위 확정을 논의함.',
      '일정 지연 및 초기 예산 초과 가능성.',
      'A안: 외주 개발 진행 (500만원)',
      'B안: 내부 인력 진행',
      'C안: 다음 분기 재검토',
      '논의중', '상', '', '', '', now, now, 'FALSE'
    ],
    [
      'ISSUE_102', '하반기 마케팅 예산 집행 건', '마케팅', '성춘향', '이몽룡', '홍길동',
      '7월 20일', '300만원',
      '온라인 브랜드 홍보 강화를 위한 광고 매체 선정 및 집행 계획안.',
      '타깃 고객 전환율 미달 가능성.',
      'A: SNS 광고 집행', 'B: 검색 광고 집행', 'C: 혼합 진행',
      '승인대기', '상', '', '', '', now, now, 'FALSE'
    ],
    [
      'ISSUE_103', '사무실 장비 교체 건', '재무', '이몽룡', '이몽룡', '홍길동',
      '8월 1일', '200만원',
      '노후 개발 장비 및 모니터 순차 교체.',
      '자금 집행 우선순위 조율 필요.',
      'A: 일시 구입', 'B: 렌탈 계약', 'C: 보류',
      '승인', '중', '', '', '', now, now, 'FALSE'
    ]
  ];

  initialIssues.forEach(function(row) { issuesSheet.appendRow(row); });

  var logsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LOGS);
  logsSheet.appendRow(['LOG_1', 'ISSUE_101', '홍길동', '-', '논의중', '최초 안건 작성', now]);
  logsSheet.appendRow(['LOG_2', 'ISSUE_102', '성춘향', '논의중', '승인대기', '최종 승인 투표 결의 개시', now]);
  logsSheet.appendRow(['LOG_3', 'ISSUE_103', '이몽룡', '승인대기', '승인', '최종결정권자 최종 승인 완료', now]);

  var commentsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.COMMENTS);
  commentsSheet.appendRow(['COM_1', 'ISSUE_101', '이몽룡', 'A안 외주 개발 업체의 포트폴리오 확인이 필요합니다.', now]);
  commentsSheet.appendRow(['COM_2', 'ISSUE_102', '성춘향', 'SNS 광고 집행 시 효율 측정이 용이합니다.', now]);

  var votesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.VOTES);
  votesSheet.appendRow(['VOT_1', 'ISSUE_102', '홍길동', '찬성', 'SNS 광고 추진에 동의합니다.', now]);
  votesSheet.appendRow(['VOT_2', 'ISSUE_102', '성춘향', '찬성', '타깃층 설정이 명확합니다.', now]);
}

function getDashboardData() {
  var ss = getDb();
  var issues = getIssuesData(ss);
  var logs = getLogsData(ss);
  var members = getMembersData(ss);

  var total = issues.length;
  var pendingApproval = 0, onHold = 0, urgent = 0, deadlineImminent = 0;

  issues.forEach(function(issue) {
    if (issue.status === '승인대기') pendingApproval++;
    if (issue.status === '보류') onHold++;
    if (issue.priority === '상') urgent++;
    if (issue.deadline && issue.deadline.indexOf("7월") !== -1 && issue.status !== '완료') {
      deadlineImminent++;
    }
  });

  return {
    summary: {
      total: total,
      pendingApproval: pendingApproval,
      onHold: onHold,
      urgent: urgent,
      deadlineImminent: deadlineImminent
    },
    recentLogs: logs.slice(0, 7),
    members: members
  };
}

function getIssues(filters) {
  var ss = getDb();
  var issues = getIssuesData(ss);

  if (!filters) return issues;

  return issues.filter(function(item) {
    if (String(item.deleted).toUpperCase() === 'TRUE') return false;

    if (filters.search && filters.search.trim() !== '') {
      var q = filters.search.toLowerCase().trim();
      var matchTitle = (item.title || '').toLowerCase().indexOf(q) !== -1;
      var matchProposer = (item.proposer || '').toLowerCase().indexOf(q) !== -1;
      var matchSummary = (item.summary || '').toLowerCase().indexOf(q) !== -1;
      if (!matchTitle && !matchProposer && !matchSummary) return false;
    }

    if (filters.tab === 'approved') {
      // 통과(결재 완료) 안건 탭: 승인, 실행중, 완료
      if (['승인', '실행중', '완료'].indexOf(item.status) === -1) return false;
    }

    if (filters.status && filters.status !== '전체' && item.status !== filters.status) return false;
    if (filters.category && filters.category !== '전체' && item.category !== filters.category) return false;
    if (filters.priority && filters.priority !== '전체' && item.priority !== filters.priority) return false;

    return true;
  });
}

function getIssueDetail(issueId) {
  var ss = getDb();
  var issues = getIssuesData(ss);
  var targetIssue = null;

  for (var i = 0; i < issues.length; i++) {
    if (issues[i].id === issueId) {
      targetIssue = issues[i];
      break;
    }
  }

  if (!targetIssue) throw new Error("존재하지 않거나 삭제된 안건입니다.");

  var comments = getCommentsData(ss, issueId);
  var votes = getVotesData(ss, issueId);
  var logs = getLogsData(ss, issueId);

  return {
    issue: targetIssue,
    comments: comments,
    votes: votes,
    logs: logs
  };
}

function createIssue(data) {
  if (!data.title || data.title.trim() === '') throw new Error("제목은 필수 입력 사항입니다.");
  if (!data.proposer || data.proposer.trim() === '') throw new Error("제안자는 필수 입력 사항입니다.");
  if (!data.decisionOwner || data.decisionOwner.trim() === '') throw new Error("최종결정자는 필수 입력 사항입니다.");
  if (!data.summary || data.summary.trim() === '') throw new Error("핵심내용은 필수 입력 사항입니다.");

  var ss = getDb();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ISSUES);
  var newId = "ISSUE_" + Date.now();
  var now = new Date().toISOString();

  var newRow = [
    newId, data.title.trim(), data.category || '기타', data.proposer.trim(), data.decisionOwner.trim(),
    data.executor || '', data.deadline || '', data.estimatedCost || '', data.summary.trim(),
    data.risk || '', data.optionA || '', data.optionB || '', data.optionC || '',
    '제안', data.priority || '중', '', '', '', now, now, 'FALSE'
  ];

  sheet.appendRow(newRow);
  addLog(ss, newId, data.proposer.trim(), '-', '제안', '최초 안건 작성');
  return { success: true, issueId: newId };
}

function addComment(issueId, author, text) {
  if (!author || author.trim() === '') throw new Error("작성자 이름을 입력해주세요.");
  if (!text || text.trim() === '') throw new Error("의견 내용을 입력해주세요.");

  var ss = getDb();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.COMMENTS);
  var now = new Date().toISOString();
  var commentId = "COM_" + Date.now();

  sheet.appendRow([commentId, issueId, author.trim(), text.trim(), now]);
  return { success: true };
}

function castVote(issueId, voter, value, reason) {
  if (!voter || voter.trim() === '') throw new Error("투표자 이름을 입력해주세요.");
  if (!value || ['찬성', '반대', '보류'].indexOf(value) === -1) throw new Error("올바른 투표 값을 선택해주세요.");

  var ss = getDb();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.VOTES);
  var data = sheet.getDataRange().getValues();
  var now = new Date().toISOString();

  var existingRowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === issueId && data[i][2] === voter.trim()) {
      existingRowIndex = i + 1;
      break;
    }
  }

  if (existingRowIndex > 0) {
    sheet.getRange(existingRowIndex, 4).setValue(value);
    sheet.getRange(existingRowIndex, 5).setValue(reason || '');
    sheet.getRange(existingRowIndex, 6).setValue(now);
  } else {
    var voteId = "VOT_" + Date.now();
    sheet.appendRow([voteId, issueId, voter.trim(), value, reason || '', now]);
  }

  return { success: true };
}

function updateIssueStatus(issueId, actor, toStatus, holdReason, reviewDate) {
  if (!actor || actor.trim() === '') throw new Error("변경 담당자(이름)를 입력해주세요.");

  var ss = getDb();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ISSUES);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var idIdx = headers.indexOf('id');
  var statusIdx = headers.indexOf('status');
  var holdReasonIdx = headers.indexOf('holdReason');
  var reviewDateIdx = headers.indexOf('reviewDate');
  var updatedAtIdx = headers.indexOf('updatedAt');

  var targetRowIndex = -1;
  var currentStatus = '';

  for (var i = 1; i < data.length; i++) {
    if (data[i][idIdx] === issueId) {
      targetRowIndex = i + 1;
      currentStatus = data[i][statusIdx];
      break;
    }
  }

  if (targetRowIndex === -1) throw new Error("안건을 찾을 수 없습니다.");

  if ((toStatus === '실행중' || toStatus === '완료') && currentStatus !== '승인' && currentStatus !== '실행중') {
    throw new Error("최종 승인(승인 상태)을 거치지 않은 안건은 '실행중' 또는 '완료' 상태로 변경할 수 없습니다.");
  }

  if (toStatus === '보류') {
    if (!holdReason || holdReason.trim() === '') throw new Error("보류 사유를 입력해야 합니다.");
    if (!reviewDate || reviewDate.trim() === '') throw new Error("재검토 일자를 설정해야 합니다.");
  }

  var now = new Date().toISOString();
  sheet.getRange(targetRowIndex, statusIdx + 1).setValue(toStatus);
  sheet.getRange(targetRowIndex, updatedAtIdx + 1).setValue(now);

  if (toStatus === '보류') {
    sheet.getRange(targetRowIndex, holdReasonIdx + 1).setValue(holdReason.trim());
    sheet.getRange(targetRowIndex, reviewDateIdx + 1).setValue(reviewDate.trim());
  }

  var logReason = "상태 변경: " + currentStatus + " -> " + toStatus;
  if (toStatus === '보류') {
    logReason += " (보류사유: " + holdReason + ", 재검토일: " + reviewDate + ")";
  }

  addLog(ss, issueId, actor.trim(), currentStatus, toStatus, logReason);
  return { success: true };
}

function deleteIssue(issueId, actor) {
  var ss = getDb();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ISSUES);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var idIdx = headers.indexOf('id');
  var deletedIdx = headers.indexOf('deleted');
  var statusIdx = headers.indexOf('status');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idIdx] === issueId) {
      sheet.getRange(i + 1, deletedIdx + 1).setValue('TRUE');
      addLog(ss, issueId, actor || '관리자', data[i][statusIdx], '삭제', '안건 삭제 처리');
      return { success: true };
    }
  }

  throw new Error("삭제할 안건을 찾지 못했습니다.");
}

function generateMeetingDoc(issueId) {
  var ss = getDb();
  var issueDetail = getIssueDetail(issueId);
  var issue = issueDetail.issue;

  var allowedStatuses = ['승인', '보류', '거절'];
  if (allowedStatuses.indexOf(issue.status) === -1) {
    throw new Error("회의록은 '승인', '보류', '거절' 상태에서만 생성 가능합니다.");
  }

  var dateStr = new Date().toISOString().split('T')[0].replace(/-/g, '');
  var docTitle = "Decision_" + issue.title + "_" + dateStr;

  var doc = DocumentApp.create(docTitle);
  var body = doc.getBody();
  body.clear();

  var titlePara = body.appendParagraph("의사결정 회의록 / 결정문");
  titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  titlePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph("문서 생성일: " + new Date().toLocaleString("ko-KR"));
  body.appendHorizontalRule();

  body.appendParagraph("1. 안건 개요").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  var summaryTableData = [
    ["안건 제목", issue.title], ["카테고리", issue.category],
    ["제안자", issue.proposer], ["최종결정자", issue.decisionOwner],
    ["실행담당자", issue.executor || "미정"], ["최종 상태", issue.status],
    ["우선순위", issue.priority], ["마감일", issue.deadline || "없음"],
    ["예상 비용", issue.estimatedCost || "없음"]
  ];

  if (issue.status === '보류') {
    summaryTableData.push(["보류 사유", issue.holdReason || ""]);
    summaryTableData.push(["재검토 일자", issue.reviewDate || ""]);
  }

  var table = body.appendTable(summaryTableData);
  table.setBorderColor("#cbd5e1");

  body.appendParagraph("\n2. 핵심 내용 및 리스크 분석").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph("■ 핵심 내용:").setBold(true);
  body.appendParagraph(issue.summary || "내용 없음");
  body.appendParagraph("\n■ 예상 리스크:").setBold(true);
  body.appendParagraph(issue.risk || "없음");

  body.appendParagraph("\n3. 비교 선택지").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph("· 선택지 A: " + (issue.optionA || "없음"));
  body.appendParagraph("· 선택지 B: " + (issue.optionB || "없음"));
  body.appendParagraph("· 선택지 C: " + (issue.optionC || "없음"));

  body.appendParagraph("\n4. 투표 및 의견 기록").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  var approveCount = 0, rejectCount = 0, holdCount = 0;
  issueDetail.votes.forEach(function(v) {
    if (v.value === '찬성') approveCount++;
    if (v.value === '반대') rejectCount++;
    if (v.value === '보류') holdCount++;
  });
  body.appendParagraph("■ 투표 집계: 찬성 " + approveCount + "표 / 반대 " + rejectCount + "표 / 보류 " + holdCount + "표").setBold(true);

  if (issueDetail.votes.length > 0) {
    var voteRows = [["투표자", "선택", "사유", "투표일시"]];
    issueDetail.votes.forEach(function(v) { voteRows.push([v.voter, v.value, v.reason || "-", v.createdAt]); });
    body.appendTable(voteRows);
  }

  body.appendParagraph("\n■ 의견 목록:").setBold(true);
  if (issueDetail.comments.length > 0) {
    issueDetail.comments.forEach(function(c) {
      body.appendParagraph("· [" + c.author + "] (" + c.createdAt + "): " + c.text);
    });
  }

  body.appendParagraph("\n5. 의사결정 진행 이력").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  if (issueDetail.logs.length > 0) {
    var logRows = [["일시", "작성자/행동자", "이전상태 -> 변경상태", "사유/설명"]];
    issueDetail.logs.forEach(function(l) {
      logRows.push([l.createdAt, l.actor, l.fromStatus + " -> " + l.toStatus, l.reason || "-"]);
    });
    body.appendTable(logRows);
  }

  doc.saveAndClose();
  var docUrl = doc.getUrl();

  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ISSUES);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idIdx = headers.indexOf('id');
  var docUrlIdx = headers.indexOf('docUrl');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idIdx] === issueId) {
      sheet.getRange(i + 1, docUrlIdx + 1).setValue(docUrl);
      break;
    }
  }

  return { success: true, docUrl: docUrl };
}

// Data helpers
function getIssuesData(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ISSUES);
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  var headers = values[0], results = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i], obj = {};
    headers.forEach(function(h, idx) { obj[h] = row[idx]; });
    results.push(obj);
  }
  return results.reverse();
}

function getCommentsData(ss, issueId) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.COMMENTS);
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  var headers = values[0], issueIdIdx = headers.indexOf('issueId'), results = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i][issueIdIdx] === issueId) {
      var obj = {};
      headers.forEach(function(h, idx) { obj[h] = values[i][idx]; });
      results.push(obj);
    }
  }
  return results;
}

function getVotesData(ss, issueId) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.VOTES);
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  var headers = values[0], issueIdIdx = headers.indexOf('issueId'), results = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i][issueIdIdx] === issueId) {
      var obj = {};
      headers.forEach(function(h, idx) { obj[h] = values[i][idx]; });
      results.push(obj);
    }
  }
  return results;
}

function getLogsData(ss, issueId) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LOGS);
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  var headers = values[0], issueIdIdx = headers.indexOf('issueId'), results = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!issueId || row[issueIdIdx] === issueId) {
      var obj = {};
      headers.forEach(function(h, idx) { obj[h] = row[idx]; });
      results.push(obj);
    }
  }
  return results.reverse();
}

function getMembersData(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MEMBERS);
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  var headers = values[0], results = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i], obj = {};
    headers.forEach(function(h, idx) { obj[h] = row[idx]; });
    results.push(obj);
  }
  return results;
}

function addLog(ss, issueId, actor, fromStatus, toStatus, reason) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LOGS);
  var logId = "LOG_" + Date.now();
  var now = new Date().toISOString();
  sheet.appendRow([logId, issueId, actor, fromStatus || '-', toStatus, reason || '', now]);
}
