// 処理終わるまで変更しないこと。
var FOLDNAME = "<INPUT FOLDER NAME>"
var DOCID = "<INPUT DOC ID>"

/*
===================================
 1.翻訳処理
=================================== 
*/
function startTranslateDocument() {
  translateDocument(DOCID, 0);
}

function retryTranslateDocument() {
  var skipCounter = parseInt(PropertiesService.getScriptProperties().getProperty("skipkey"))
  deleteTriggerSettings()
  translateDocument(DOCID, skipCounter)
}

function translateDocument(d_id, skipCount) {
  var startTime = new Date();
  d = DocumentApp.openById(d_id)
  Logger.log(d);
  name = d.getName()
  index = 0
  paragraphs = d.getBody().getParagraphs()
  flg = true;
  paragraphs.forEach(function (paragraph, index) {
    if (PropertiesService.getScriptProperties().getProperty("skipkey")) {
      Logger.log("Trigger Defined");
      return;
    } else {
      // 処理時間によりリトライトリガーを作成する。
      if (retrySetting(this.startTime, index, "retryTranslateDocument")) {
        flg = false
        Logger.log("RetrySet");
        return;
      }
    }

    text = paragraph.getText()
    // v8エンジン対応
    if (text.indexOf("[Formula]") === -1 && paragraph.getHeading() == "NORMAL" && index >= this.skipCount && flg) {
      Logger.log("Processing")
      var j = LanguageApp.translate(text, 'en', 'ja')
      paragraph.appendText("\n" + j)
      Utilities.sleep(1200)
    } else {
      Logger.log("not translate")
    }
  }, { 'skipCount': skipCount, 'startTime': startTime });
}

/*
===================================
 2.OCR一括処理
=================================== 
*/
function startExecute() {
  mainProcess(0);
}

// 万が一途中で止まった場合
function setdocpro() {
  PropertiesService.getScriptProperties().setProperty("skipkey", 100);
}


function doRetry() {
  var skipCounter = parseInt(PropertiesService.getScriptProperties().getProperty("skipkey"))
  deleteTriggerSettings()
  mainProcess(skipCounter)
}

function deleteTriggerSettings() {
  var triggerId = PropertiesService.getScriptProperties().getProperty("tid");
  ScriptApp.getProjectTriggers().filter(function (trigger) {
    return trigger.getUniqueId() == triggerId;
  }).forEach(function (trigger) {
    ScriptApp.deleteTrigger(trigger);
  });
  PropertiesService.getScriptProperties().deleteAllProperties()
}

// リトライ用の処理
function retrySetting(startTime, index, funcionName) {
  Logger.log("Retry Setting")
  RETRY_INTERVAL = 240
  if (parseInt((new Date() - startTime)) > RETRY_INTERVAL * 1000) {
    var dt = new Date()
    dt.setMinutes(dt.getMinutes() + 2)
    var triggerId = ScriptApp.newTrigger(funcionName).timeBased().at(dt).create().getUniqueId();
    PropertiesService.getScriptProperties().setProperty("skipkey", index);
    PropertiesService.getScriptProperties().setProperty("tid", triggerId)
    return true;
  }
  return false;
}

// 2分後にOCR処理されたファイルマージ処理を動かす。
function secondSetting() {
  Logger.log("")
  var dt = new Date()
  dt.setMinutes(dt.getMinutes() + 2)
  var triggerId = ScriptApp.newTrigger("processSecond").timeBased().at(dt).create().getUniqueId();
  PropertiesService.getScriptProperties().setProperty("tid", triggerId)

}

// ターゲットファイルを取得する。
function getTargetFiles() {
  var files = DriveApp.getFoldersByName(FOLDNAME)
  Logger.log(files)

  if (!files) {
    return null
  }

  var folderId = ""
  var iter
  var retFiles = []
  while (files.hasNext()) {
    folderId = files.next().getId()
  }

  // Idがない場合は空配列を返すよう修正   
  if (folderId) {
    // Iteratorではなく、リストに変換
    iter = DriveApp.getFolderById(folderId).getFiles()
    while (iter.hasNext()) {
      retFiles.push(iter.next())
    }
  }
  Logger.log(retFiles)
  return retFiles
}

// OCR main処理
function mainProcess(skipCount) {
  var startTime = new Date();
  var ocrFileList = []
  var fileList = []
  var list = []
  var proceedFlg = true;

  ocrFileList = getTargetFiles();
  ocrFileList.forEach(function (ocrfile, index) {
    Logger.log(ocrfile.getName())

    // 翻訳の処理を追加
    if (ocrfile.getName() !== ".DS_Store" && index >= this.skipCount) {
      Logger.log("!")
      //定義があれば、やめる。
      if (PropertiesService.getScriptProperties().getProperty("skipkey")) {
        return;
      } else {
        // 処理時間によりリトライトリガーを作成する。
        if (retrySetting(this.startTime, index, "doRetry")) {
          proceedFlg = false;
          return;
        }
      }
      doocr(ocrfile)
    }
  }, { 'skipCount': skipCount, 'startTime': startTime });

  if (proceedFlg) {
    Logger.log("Create trigger secondSetting()")
    secondSetting();
  }

　　/*
     Private Method化: OCRの処理を実行する。
    */
  function doocr(file) {
    mediaData = file.getBlob();
    var resource = {
      title: mediaData.getName(),
      mimeType: mediaData.getContentType()
    };
    // OCRの設定
    var optionalArgs = {
      ocr: true,
      ocrLanguage: 'ja'
    };
    // Google Driveにファイル追加
    Drive.Files.insert(resource, mediaData, optionalArgs);
  }
}

function processSecond() {
  // とりあえずトリガー設定は削除する。
  deleteTriggerSettings()
  var doc = DocumentApp.create('New_' + FOLDNAME)
  var body = doc.getBody()
  var filelist = []
  var docList = []
  var deleteList = []
  ocrFileList = getTargetFiles();
  ocrFileList.forEach(function (ocrFile) {
    filelist.push(ocrFile.getName())
  })
  // 前処理でルートフォルダにOCR処理済のdocをinsertしているため
  t = DriveApp.getRootFolder().getFiles()
  var ids = []
  while (t.hasNext()) {
    var file = t.next()
    // jpgとpngで検索
    var fileTypeList = [file.getName() + ".png", file.getName() + ".jpg"]

    for (fIndex in fileTypeList) {
      index = filelist.indexOf(fileTypeList[fIndex])
      if (index !== -1) {
        obj = makeobj(file)
        if (obj) {
          docList.push(obj)
          deleteList.push(file.getId())
        }
      }
    }
  }

  sList = docList.sort(function (a, b) {
    if (a['id'] > b['id']) {
      return 1
    } else if (a['id'] < b['id']) {
      return -1
    }
    return 0
  })
  body.appendParagraph(FOLDNAME).setHeading(DocumentApp.ParagraphHeading.HEADING1)
  body.appendParagraph("heading3").setHeading(DocumentApp.ParagraphHeading.HEADING3)
  docList.forEach(function (e) {
    body.appendParagraph(e["header"]).setHeading(DocumentApp.ParagraphHeading.HEADING2)
    body.appendParagraph(e["content"]).setHeading(DocumentApp.ParagraphHeading.NORMAL)
  })

  // ヘッダ設定の調整
  setHeadingAttribute(doc.getId())
  // 不要ファイルの削除  
  deleteList.forEach(function (fileId) {
    DriveApp.removeFile(DriveApp.getFileById(fileId))
  })

  function makeobj(file) {
    var obj = {}
    var content = DocumentApp.openById(file.getId()).getBody().getText()
    obj['id'] = file.getName()
    obj['header'] = file.getBlob().getName()
    obj['content'] = content

    if (content) {
      return obj
    }
    return null
  }

}



/*
結合する。
*/
function concatdoc(condition) {

  var doc = DocumentApp.create('ConcatDoc' + condition);
  var body = doc.getBody()
  var files = DriveApp.searchFiles('title contains "' + condition + '"')
  var text = ""
  var list = []

  while (files.hasNext()) {
    var obj = {}
    var file = files.next();
    id = file.getId()
    contentType = file.getBlob().getContentType()
    Logger.log(file.getName() + ":" + contentType)
    if (contentType === "application/pdf" && file.getName() !== ".DS_Store") {
      name = file.getBlob().getName()
      var a = name.replace(condition, "").replace("_", "").replace(".pdf", "")
      var header = file.getBlob().getName()
      var content = DocumentApp.openById(id).getBody().getText()
      obj['id'] = a
      obj['header'] = header
      obj['content'] = content
      list.push(obj)
    }
  }
  s_list = list.sort(function (a, b) {
    if (a['id'] > b['id']) {
      return 1
    } else if (a['id'] < b['id']) {
      return -1
    }
    return 0
  })
  list.forEach(function (e) {
    body.appendParagraph(e["header"]).setHeading(DocumentApp.ParagraphHeading.HEADING2)
    body.appendParagraph(e["content"]).setHeading(DocumentApp.ParagraphHeading.NORMAL)
  })
}


/*
見出しの調整
*/
function midashi() {
  const style = {}

  DocumentApp.ParagraphHeading.HEADING1
  var files = DriveApp.searchFiles('title = "ConCatDocr_at_chapter03"')
  searchPattern = "r_at_chapter"
  while (files.hasNext()) {
    var file = files.next()
    var id = file.getId()
    docs = DocumentApp.openById(id)
    var range = docs.getBody().findText(searchPattern)
    while (range) {
      range.getElement().
        Logger.log(range.getElement().asText().getText())
      range = docs.getBody().findText(searchPattern)
    }
  }
}


function setHeading() {
  setHeadingAttribute(DOCID)
}

// 見出しを設定する。
function setHeadingAttribute(doc_id) {

  var body = DocumentApp.openById(doc_id).getBody()

  var ps = body.getParagraphs()

  //  var ps = body.getParagraphs
  var normal_styles = {}
  var heading1_styles = {}
  var heading2_styles = {}
  var heading3_styles = {}

  margin = 72
  body.setMarginBottom(margin)
  body.setMarginLeft(margin)
  body.setMarginRight(margin)
  body.setMarginTop(margin)

  normal_styles[DocumentApp.Attribute.FONT_SIZE] = 11
  normal_styles[DocumentApp.Attribute.FONT_FAMILY] = "Roboto"
  normal_styles[DocumentApp.Attribute.LINE_SPACING] = 1.15
  normal_styles[DocumentApp.Attribute.SPACING_BEFORE] = 6.0
  normal_styles[DocumentApp.Attribute.SPACING_AFTER] = 6.0

  heading1_styles[DocumentApp.Attribute.FONT_SIZE] = 18
  heading1_styles[DocumentApp.Attribute.FONT_FAMILY] = "Roboto"
  heading1_styles[DocumentApp.Attribute.LINE_SPACING] = 1.15
  heading1_styles[DocumentApp.Attribute.BOLD] = false
  heading1_styles[DocumentApp.Attribute.SPACING_BEFORE] = 6.0
  heading1_styles[DocumentApp.Attribute.SPACING_AFTER] = 6.0

  heading2_styles[DocumentApp.Attribute.FONT_SIZE] = 14
  heading2_styles[DocumentApp.Attribute.FONT_FAMILY] = "Roboto"
  heading2_styles[DocumentApp.Attribute.LINE_SPACING] = 1.15
  heading2_styles[DocumentApp.Attribute.BOLD] = true
  heading2_styles[DocumentApp.Attribute.SPACING_BEFORE] = 6.0
  heading2_styles[DocumentApp.Attribute.SPACING_AFTER] = 6.0

  heading3_styles[DocumentApp.Attribute.FONT_SIZE] = 12
  heading3_styles[DocumentApp.Attribute.FONT_FAMILY] = "Roboto"
  heading3_styles[DocumentApp.Attribute.LINE_SPACING] = 1.15
  heading3_styles[DocumentApp.Attribute.BOLD] = false
  heading3_styles[DocumentApp.Attribute.ITALIC] = true
  heading3_styles[DocumentApp.Attribute.SPACING_BEFORE] = 6.0
  heading3_styles[DocumentApp.Attribute.SPACING_AFTER] = 6.0

  for (p in ps) {
    Logger.log(ps[p].getText() + ":" + ps[p].getHeading())
    if (ps[p].getHeading() == "NORMAL") {
      ps[p].setAttributes(normal_styles)
    } else if (ps[p].getHeading() == "HEADING1") {
      ps[p].setAttributes(heading1_styles)
    } else if (ps[p].getHeading() == "HEADING2") {
      ps[p].setAttributes(heading2_styles)
    } else if (ps[p].getHeading() == "HEADING3") {
      ps[p].setAttributes(heading3_styles)
    }
  }
}

