/**
 * submit-v2.html 신청폼 수신용 Google Apps Script
 *
 * 설정 방법
 * 1. 접수 내역을 쌓을 Google 스프레드시트를 새로 만들고, 메뉴 [확장 프로그램 → Apps Script]를 연다.
 * 2. 이 파일 내용을 그대로 붙여넣고, 아래 NOTIFY_EMAIL 을 알림 받을 주소로 바꾼다 (여러 명은 쉼표로 구분).
 * 3. [배포 → 새 배포 → 유형: 웹 앱]
 *      - 다음 사용자 인증 정보로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 *    배포 후 나오는 웹 앱 URL(https://script.google.com/macros/s/.../exec)을 복사한다.
 * 4. submit-v2.html 하단 스크립트의 ENDPOINT = "" 에 그 URL을 넣는다.
 *
 * 접수될 때마다: 스프레드시트에 한 줄 추가, 첨부파일은 내 드라이브 "전년도 증권 제출" 폴더에 저장,
 * NOTIFY_EMAIL 로 알림 메일 발송.
 */

var NOTIFY_EMAIL = "kta7491@naver.com,jykim@ncrm.biz";
var FOLDER_NAME = "전년도 증권 제출";
var SHEET_NAME = "전년도 증권 제출";
var MAX_BYTES = 20 * 1024 * 1024;

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var list = d.files || (d.file ? [d.file] : []);
    if (!d.hotel || !d.phone || !d.email || !d.agree || !list.length) {
      return reply({ ok: false, error: "missing fields" });
    }

    var blobs = list.map(function (f) { return { f: f, bytes: Utilities.base64Decode(f.data) }; });
    var total = blobs.reduce(function (sum, b) { return sum + b.bytes.length; }, 0);
    if (total > MAX_BYTES) return reply({ ok: false, error: "file too large" });

    var now = new Date();
    var stamp = Utilities.formatDate(now, "Asia/Seoul", "yyyyMMdd-HHmmss");
    var folder = getFolder();
    var urls = blobs.map(function (b) {
      var name = stamp + "_" + d.hotel + "_" + b.f.name;
      return folder.createFile(Utilities.newBlob(b.bytes, b.f.type, name)).getUrl();
    });

    getSheet().appendRow([
      Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss"),
      d.hotel, d.manager || "", "'" + d.phone, d.email, "동의", urls.join("\n")
    ]);

    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: "[전년도 증권 제출] " + d.hotel,
      body: [
        "호텔명: " + d.hotel,
        "담당자명: " + (d.manager || "-"),
        "전화번호: " + d.phone,
        "E-mail: " + d.email,
        "첨부파일 (" + urls.length + "개):\n" + urls.join("\n")
      ].join("\n"),
      replyTo: d.email
    });

    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  }
}

function getFolder() {
  var it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["접수일시", "호텔명", "담당자명", "전화번호", "E-mail", "개인정보 동의", "첨부파일"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
