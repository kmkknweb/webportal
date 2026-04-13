// ==========================================
// TEMPLATES — เพิ่ม template ใหม่ตรงนี้ครับ
// ==========================================
var TEMPLATES = [
  {
    id: '1c5UJ0wEZTO9NqbHNMNnP4AwWKBsXq5LTLMNvbkTQQb4',
    name: 'ใบประกาศเกียรติคุณ'
  },
  // เพิ่ม template ใหม่แบบนี้ครับ:
  // { id: 'SLIDES_ID_ของ_template_ใหม่', name: 'ใบอนุโมทนาบัตร' },
  // { id: 'SLIDES_ID_ของ_template_ใหม่', name: 'ใบประกาศนียบัตร' },
];

// โฟลเดอร์เก็บ PDF ที่สร้างออกมา
var FOLDER_ID = '1Pr9NS6UQRWGD5bKKspQp99iv94qg3y2E';

// ==========================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Certificate')
    .setTitle('ระบบสร้างใบประกาศ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ส่ง template list ไปให้หน้าเว็บ
function getTemplates() {
  return TEMPLATES.map(function(t) {
    return { id: t.id, name: t.name };
  });
}

function generateCertificate(templateId, name, event, font, fontSize, fontColor) {
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var templateFile = DriveApp.getFileById(templateId);

  // copy template
  var copyFile = templateFile.makeCopy("temp_" + name, folder);
  var copyId = copyFile.getId();

  var presentation = SlidesApp.openById(copyId);
  var slide = presentation.getSlides()[0];

  // แทนชื่อก่อน
  slide.replaceAllText("{{NAME}}", name);

  // จัดการชื่อ + event
  var shapes = slide.getShapes();
  for (var i = 0; i < shapes.length; i++) {
    try {
      var textRange = shapes[i].getText();
      var rawText = textRange.asString();

      // EVENT
      if (rawText.indexOf("{{EVENT}}") > -1) {
        textRange.setText(event);
      }

      // NAME: ปรับ style เฉพาะกล่องชื่อ
      var finalText = textRange.asString().trim();
      if (finalText === name) {
        var style = textRange.getTextStyle();
        style.setFontFamily(font);
        style.setFontSize(Number(fontSize));
        style.setForegroundColor(fontColor);
        style.setBold(true);
      }
    } catch (e) {}
  }

  presentation.saveAndClose();

  var pdfBlob = DriveApp.getFileById(copyId).getBlob().getAs(MimeType.PDF);
  var fileName = "certificate_" + name + ".pdf";
  var pdfFile = folder.createFile(pdfBlob).setName(fileName);

  pdfFile.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  DriveApp.getFileById(copyId).setTrashed(true);

  Utilities.sleep(1500);

  return "https://drive.google.com/file/d/" + pdfFile.getId() + "/view?usp=sharing";
}
