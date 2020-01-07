var FOLDNAME = "<INPUT FOLDER NAME>"
/*
===================================
 1.翻訳処理
=================================== 
*/
function startTranslateDocument(){
  translateDocument(0); 
}

// Not tested.
function retryTranslateDocument(){
    var skipCounter = parseInt(PropertiesService.getScriptProperties().getProperty("skipkey")) 
    deleteTriggerSettings()
    mainProcess(skipCounter)
}

// 
function translateDocument(skipCount){
  var startTime = new Date();
  id = "<INPUT DOC ID>"
  d = DocumentApp.openById(id)
  name = d.getName()
  index = 0
  paragraphs = d.getBody().getParagraphs()  
  paragraphs.forEach(function(paragraph,index){    
    if (retrySetting(this.startTime,index,"retryTranslateDocument") ){ 
       return; 
    }
    text = paragraph.getText()
    if(text.indexOf("[Formula]") == -1 && paragraph.getHeading() == "Normal" && index >= this.skipCount){
      var j = LanguageApp.translate(text, 'en', 'ja')
      paragraph.appendText("\n"+j)
      Utilities.sleep(1500)
    }
  },{'skipCount':skipCount,'startTime':startTime});
}

/*
===================================
 2.OCR一括処理
=================================== 
*/
function startExecute(){
    mainProcess(0);
}

// 万が一途中で止まった場合
function setdocpro(){
    PropertiesService.getScriptProperties().setProperty("skipkey", 100);  
}


function doRetry(){
    // https://kido0617.github.io/js/2017-02-13-gas-6-minutes/
    var skipCounter = parseInt(PropertiesService.getScriptProperties().getProperty("skipkey")) 
    deleteTriggerSettings()
    mainProcess(skipCounter)
}

function deleteTriggerSettings(){
    var triggerId = PropertiesService.getScriptProperties().getProperty("tid");
    ScriptApp.getProjectTriggers().filter(function(trigger){
      return trigger.getUniqueId() == triggerId;
    }).forEach(function(trigger) {
      ScriptApp.deleteTrigger(trigger);
    });
    PropertiesService.getScriptProperties().deleteAllProperties()
}  

// リトライ用の処理
function retrySetting(startTime, index,funcionName){
    if( parseInt((new Date() - startTime)) > 300 * 1000){
         var dt = new Date()
         dt.setMinutes(dt.getMinutes() + 1)
         var triggerId = ScriptApp.newTrigger(funcionName).timeBased().at(dt).create().getUniqueId();
         PropertiesService.getScriptProperties().setProperty("skipkey", index);  
         PropertiesService.getScriptProperties().setProperty("tid", triggerId)
         return true;
    } 
    return false;
}

// 3分後くらいに次の処理をしこむ
function secondSetting(){
    var dt = new Date()
    dt.setMinutes(dt.getMinutes() + 3)
    var triggerId = ScriptApp.newTrigger("processSecond").timeBased().at(dt).create().getUniqueId();  
    PropertiesService.getScriptProperties().setProperty("tid", triggerId)
}

// ターゲットファイルを取得する。
function getTargetFiles(){
   var files = DriveApp.getFoldersByName(FOLDNAME)
   var folderId = ""
   var iter
   var retFiles = []
   while (files.hasNext()){
     folderId = files.next().getId()
   }
   // Iteratorではなく、リストに変換
   iter = DriveApp.getFolderById(folderId).getFiles()
   while(iter.hasNext()){
     retFiles.push(iter.next())
   }
   return retFiles
}

// OCR main処理
function mainProcess(skipCounter){
   var startTime = new Date();
   var ocrFileList = []
   var fileList = []
   var list = []

   ocrFileList = getTargetFiles();
   ocrFileList.forEach(function(ocrfile,index){
     if( ocrfile.getName() !== ".DS_Store" && index >= this.skipCount){
       // 5分過ぎれば、トリガーを作成して処理を終了する。
       if (retrySetting(this.startTime,index,"doRetry") ){ 
         return; 
       }
       doocr(ocrfile)
     } 
   },{'skipCount':skipCount, 'startTime':startTime })(); 
   secondSetting();
}

function processSecond(){
  // とりあえずトリガー設定は削除する。
  deleteTriggerSettings()
  var doc = DocumentApp.create('New_'+ FOLDNAME)
  var body = doc.getBody()
  var filelist = []
  var docList = []
  var deleteList = []
  ocrFileList = getTargetFiles();
  ocrFileList.forEach(function(ocrFile){
    filelist.push(ocrFile.getName())
  })
  // 前処理でルートフォルダにOCR処理済のdocをinsertしているため
  t = DriveApp.getRootFolder().getFiles()
  var ids = []
  while(t.hasNext()){
    var file = t.next()
     // jpgとpngで検索
    var fileTypeList = [file.getName()+".png",file.getName()+".jpg"]
     
    for( fIndex in fileTypeList ){       
      index = filelist.indexOf(searchList[fIndex])
      if(index !== -1){
        obj = makeobj(file)
        if(obj){
          docList.push(obj)
          deleteList.push(file.getId())
        }
      }
    }
  }
  
  sList = docList.sort(function(a,b){
    if(a['id'] > b['id']){
       return 1 
    } else if(a['id'] < b['id']){
       return -1
    }
    return 0
  })
  body.appendParagraph(FOLDNAME).setHeading(DocumentApp.ParagraphHeading.HEADING1)
  body.appendParagraph("heading3").setHeading(DocumentApp.ParagraphHeading.HEADING3)
  docList.forEach(function(e){
    body.appendParagraph(e["header"]).setHeading(DocumentApp.ParagraphHeading.HEADING2)
    body.appendParagraph(e["content"]).setHeading(DocumentApp.ParagraphHeading.NORMAL)
  })
  
  // ヘッダ設定の調整
  setHeadingAttribute(doc.getId())
  // 不要ファイルの削除  
  deleteList.forEach(function(fileId){
    DriveApp.removeFile(DriveApp.getFileById(fileId))
  })
}  

function makeobj(file){
  var obj = {}
  var content = DocumentApp.openById(file.getId()).getBody().getText()
  obj['id'] = file.getName()
  obj['header'] = file.getBlob().getName()
  obj['content'] = content
  
  if(content){
    return obj
  }
  return null
}


/*
OCRするためのファイルを検索する。
*/
function searchFile(condition){
  var files = DriveApp.searchFiles('title contains "'+condition + '"');
  while(files.hasNext()){
    var file = files.next();
    Logger.log(file)
    doocr(file)
  }
}

/*
OCRの処理を実行する。
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

/*
結合する。
*/
function concatdoc(condition){

  var doc = DocumentApp.create('ConcatDoc'+ condition);
  var body = doc.getBody()
  var files = DriveApp.searchFiles('title contains "'+condition + '"')
  var text = ""
  var list = []
  
  while(files.hasNext()){
    var obj = {}
    var file = files.next();
    id = file.getId()
    contentType = file.getBlob().getContentType()
    Logger.log(file.getName() + ":" + contentType)
    if( contentType === "application/pdf" && file.getName() !== ".DS_Store"){
      name = file.getBlob().getName()
      var a = name.replace(condition,"").replace("_","").replace(".pdf","")
      var header = file.getBlob().getName()
      var content = DocumentApp.openById(id).getBody().getText()
      obj['id'] = a
      obj['header'] = header
      obj['content'] = content
      list.push(obj)
    }
  }
  s_list = list.sort(function(a,b){
    if(a['id'] > b['id']){
      return 1 
    }else if(a['id'] < b['id']){
      return -1
    }
    return 0
  })
  list.forEach(function(e){
    body.appendParagraph(e["header"]).setHeading(DocumentApp.ParagraphHeading.HEADING2)
    body.appendParagraph(e["content"]).setHeading(DocumentApp.ParagraphHeading.NORMAL)
  })
}


/*
見出しの調整
*/
function midashi(){
  const style = {}
  
  DocumentApp.ParagraphHeading.HEADING1
  var files = DriveApp.searchFiles('title = "ConCatDocr_at_chapter03"')
  searchPattern = "r_at_chapter"
  while(files.hasNext()){
    var file = files.next()
    var id = file.getId()
    docs = DocumentApp.openById(id)
    var range = docs.getBody().findText(searchPattern)
    while(range){
      range.getElement().
      Logger.log(range.getElement().asText().getText())
      range = docs.getBody().findText(searchPattern)
    }
  }
}

// 見出しを設定する。
function setHeadingAttribute(doc_id){

  var body = DocumentApp.openById(doc_id).getBody()
  var ps = body.getParagraphs()
  var normal_styles = {}
  var heading1_styles = {}
  var heading2_styles = {}
  var heading3_styles = {}
  
  normal_styles[DocumentApp.Attribute.FONT_SIZE]=11
  normal_styles[DocumentApp.Attribute.FONT_FAMILY] = "Roboto"
  normal_styles[DocumentApp.Attribute.LINE_SPACING] = 1.15
  normal_styles[DocumentApp.Attribute.SPACING_BEFORE] = 6.0
  normal_styles[DocumentApp.Attribute.SPACING_AFTER] = 6.0

  heading1_styles[DocumentApp.Attribute.FONT_SIZE]=18
  heading1_styles[DocumentApp.Attribute.FONT_FAMILY] = "Roboto"
  heading1_styles[DocumentApp.Attribute.LINE_SPACING] = 1.15
  heading1_styles[DocumentApp.Attribute.BOLD] = false
  heading1_styles[DocumentApp.Attribute.SPACING_BEFORE] = 6.0
  heading1_styles[DocumentApp.Attribute.SPACING_AFTER] = 6.0  
  
  heading2_styles[DocumentApp.Attribute.FONT_SIZE]=14
  heading2_styles[DocumentApp.Attribute.FONT_FAMILY] = "Roboto"
  heading2_styles[DocumentApp.Attribute.LINE_SPACING] = 1.15
  heading2_styles[DocumentApp.Attribute.BOLD] = true
  heading2_styles[DocumentApp.Attribute.SPACING_BEFORE] = 6.0
  heading2_styles[DocumentApp.Attribute.SPACING_AFTER] = 6.0
  
  heading3_styles[DocumentApp.Attribute.FONT_SIZE]=12
  heading3_styles[DocumentApp.Attribute.FONT_FAMILY] = "Roboto"
  heading3_styles[DocumentApp.Attribute.LINE_SPACING] = 1.15
  heading3_styles[DocumentApp.Attribute.BOLD] = false
  heading3_styles[DocumentApp.Attribute.ITALIC] = true
  heading3_styles[DocumentApp.Attribute.SPACING_BEFORE] = 6.0
  heading3_styles[DocumentApp.Attribute.SPACING_AFTER] = 6.0


  for(p in ps){
    if(ps[p].getHeading()=="Normal") {
      ps[p].setAttributes(normal_styles)
    } else if (ps[p].getHeading()=="Heading 1"){
      ps[p].setAttributes(heading1_styles)
    } else if (ps[p].getHeading()=="Heading 2"){
      ps[p].setAttributes(heading2_styles)
    } else if (ps[p].getHeading() == "Heading 3"){
      ps[p].setAttributes(heading3_styles)
    } 
  }
  
  
}

