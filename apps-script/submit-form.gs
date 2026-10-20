/**
 * submit-v2.html(전년도 증권 제출) · consult-v2.html(상담 신청) 신청폼 수신용 Google Apps Script
 *
 * 설정 방법
 * 1. 접수 내역을 쌓을 Google 스프레드시트를 새로 만들고, 메뉴 [확장 프로그램 → Apps Script]를 연다.
 * 2. 이 파일 내용을 그대로 붙여넣고, 아래 NOTIFY_EMAIL 을 알림 받을 주소로 바꾼다 (여러 명은 쉼표로 구분).
 * 3. [배포 → 새 배포 → 유형: 웹 앱]
 *      - 다음 사용자 인증 정보로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 *    배포 후 나오는 웹 앱 URL(https://script.google.com/macros/s/.../exec)을 복사한다.
 * 4. 각 html 하단 스크립트의 ENDPOINT 에 그 URL을 넣는다.
 *    코드를 고친 뒤에는 [배포 관리 → 수정 → 버전: 새 버전]으로 배포해야 URL이 유지된다.
 *
 * 접수될 때마다: 스프레드시트의 해당 탭에 한 줄 추가, 첨부파일은 내 드라이브의 해당 폴더에 저장,
 * NOTIFY_EMAIL 로 알림 메일 발송.
 */

var NOTIFY_EMAIL = "kta7491@naver.com,jykim@ncrm.biz";
var MAX_BYTES = 20 * 1024 * 1024;

// 폼 종류별 설정. fields: [시트/메일에 쓸 이름, 요청 데이터 키, 필수 여부]
var FORMS = {
  submit: {
    title: "전년도 증권 제출",
    fileRequired: true,
    fields: [
      ["호텔명", "hotel", true],
      ["담당자명", "manager", false],
      ["전화번호", "phone", true],
      ["E-mail", "email", true]
    ]
  },
  consult: {
    title: "상담 신청",
    fileRequired: false,
    fields: [
      ["호텔명", "hotel", true],
      ["사업자번호", "bizno", true],
      ["담당자명", "manager", false],
      ["전화번호", "phone", true],
      ["E-mail", "email", true],
      ["전년도 보험사", "insurer", false],
      ["보험개시일", "startDate", false],
      ["전년도 보험료", "premium", false]
    ]
  }
};

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var form = FORMS[d.type || "submit"];
    if (!form) return reply({ ok: false, error: "unknown form" });

    var list = d.files || (d.file ? [d.file] : []);
    var missing = form.fields.some(function (f) { return f[2] && !d[f[1]]; });
    if (missing || !d.agree || (form.fileRequired && !list.length)) {
      return reply({ ok: false, error: "missing fields" });
    }

    var blobs = list.map(function (f) { return { f: f, bytes: Utilities.base64Decode(f.data) }; });
    var total = blobs.reduce(function (sum, b) { return sum + b.bytes.length; }, 0);
    if (total > MAX_BYTES) return reply({ ok: false, error: "file too large" });

    var now = new Date();
    var stamp = Utilities.formatDate(now, "Asia/Seoul", "yyyyMMdd-HHmmss");
    var urls = [];
    if (blobs.length) {
      var folder = getFolder(form.title);
      urls = blobs.map(function (b) {
        var name = stamp + "_" + d.hotel + "_" + b.f.name;
        return folder.createFile(Utilities.newBlob(b.bytes, b.f.type, name)).getUrl();
      });
    }

    // 앞에 ' 를 붙여 전화번호·사업자번호 등이 숫자/날짜로 바뀌지 않게 한다
    var row = [Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss")]
      .concat(form.fields.map(function (f) { return d[f[1]] ? "'" + d[f[1]] : ""; }))
      .concat(["동의", urls.join("\n")]);
    getSheet(form).appendRow(row);

    var body = form.fields.map(function (f) { return f[0] + ": " + (d[f[1]] || "-"); });
    body.push(urls.length ? "첨부파일 (" + urls.length + "개):\n" + urls.join("\n") : "첨부파일: 없음");
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: "[" + form.title + "] " + d.hotel,
      body: body.join("\n"),
      replyTo: d.email
    });

    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  }
}

function getFolder(name) {
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function getSheet(form) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(form.title);
  if (!sh) {
    sh = ss.insertSheet(form.title);
    sh.appendRow(["접수일시"].concat(form.fields.map(function (f) { return f[0]; })).concat(["개인정보 동의", "첨부파일"]));
    sh.setFrozenRows(1);
  }
  return sh;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
