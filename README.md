/////GS////////
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate().setTitle('ระบบรับสมัครนักเรียนออนไลน์ ปีการศึกษา 2563');//แก้จุดที่ 1
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('PDF')
      .addItem('สร้างไฟล์ PDF','run_pdf')
      .addToUi();
}
function run_pdf() {
  runPDFs();
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function processForm(formObject) {
  var msg = 'แจ้งมีนักเรียนมาสมัครเรียนแล้ว 1 คน ชื่อ :';  //แก้จุดที่ 2
  msg += ' \n' + formObject.first_name+" "+formObject.last_name;
  
  var url = "aaa"; //แก้จุดที่ 3
  var ss = SpreadsheetApp.openByUrl(url);
  var ws = ss.getSheetByName("aaa");//แก้จุดที่ 4
  var date = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
  ws.appendRow([formObject.first_name,
                formObject.last_name,
                formObject.iduser,
                formObject.gender,
                formObject.dateOfBirth,
                formObject.phone,
                formObject.address,
                formObject.nickname,
                formObject.class,
                formObject.plan,
                date,
                formObject.email,
                formObject.blood,
                formObject.picID 
               ]);

  sendLineNotify(msg);
//  MailApp.sendEmail("ใส่อีเมลที่ต้องการ","สมัครเข้าเรียน","มีคนสมัครเข้าเรียนแล้ว 1 คน");
//  MailApp.sendEmail(formObject.email, "การสมัครเข้าเรียน", "นักเรียนชื่อ " + formObject.first_name+" "+formObject.last_name+" สมัครเข้าเรียนระดับชั้น"+formObject.class+"ทางโรงเรียนได้รับข้อมูลการสมัครเข้าเรียนเรียบร้อยแล้ว");
}

function sendLineNotify(message) {

var token = "aaa"; //แก้จุดที่ 5
var options = {
"method": "post",
"payload": "message=" + message,
"headers": {
"Authorization": "Bearer " + token
}
};

UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}

/* สร้างไฟล์ pdf */
function runPDFs() {

const docFile = DriveApp.getFileById("aaa"); //แก้จุดที่ 6
const tempFolder = DriveApp.getFolderById("aaa"); //แก้จุดที่ 7
const pdfFolder = DriveApp.getFolderById("aaa"); //แก้จุดที่ 8
const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("aaa"); //แก้จุดที่ 9

//แก้จุดที่ 10
var patternData=["{ชื่อ}","{สกุล}","{เลขบัตรประชาชน}","{เพศ}","{วันเกิด}","{เบอร์โทร}","{ที่อยู่}",
"{ชื่อเล่น}","{ระดับชั้น}","{แผนการเรียน}","{วันที่สมัคร}", "{email}", "{กรุ๊ปเลือด}", "{รูปภาพ}"];
  
 /* กำหนดช่วงข้อมูลที่จะนำไปสร้างไฟล์ pdf */
const data = currentSheet.getRange(2, 1, currentSheet.getLastRow()-1,currentSheet.getLastColumn()).getDisplayValues();

data.forEach(row => {createPDF(row, patternData,row[0]+"-"+row[1],docFile, tempFolder, pdfFolder)});

}

/* กำหนดค่าที่จะนำมาสร้างไฟล์ pdf */

function createPDF ( data, pattern, pdfName,  docFile, tempFolder, pdfFolder) {
const tempFile = docFile.makeCopy(tempFolder);
const tempDocFile = DocumentApp.openById(tempFile.getId());
const body = tempDocFile.getBody();
var replaceTextToImage = function( body,searchText, image, width) {

    var next = body.findText(searchText);
    if (!next) return;
    var r = next.getElement();
    r.asText().setText("");
    var img = r.getParent().asParagraph().insertInlineImage(0, image);
    if (width && typeof width == "number") {
      var w = img.getWidth();
      var h = img.getHeight();
      img.setWidth(width);
      img.setHeight(width * h / w);
    }
    return next;
  };

  for (var i=0;i<data.length-1;i++){
    body.replaceText(pattern[i], data[i]);
    Logger.log(pattern[i]+" : "+data[i]);
  }
  var replaceText = pattern[data.length-1]; //ชื่อข้อความใน doc ที่ต้องการจะเปลี่ยนเป็นรูป
  var imageFileId = data[data.length-1]; //id ในdrive ตรวจสอบลำดับ คอลัมน์รูปภาพให้ดีก่อน     
  var image = DriveApp.getFileById(imageFileId).getBlob();
      next = replaceTextToImage(body, replaceText, image, 100);
tempDocFile.saveAndClose();
const pdfContentBlob = tempFile.getAs(MimeType.PDF);
pdfFolder.createFile(pdfContentBlob).setName(pdfName);
tempFolder.removeFile(tempFile);
}


/* อัปโหลดรูป */
function uploadFileToGoogleDrive(data, file) {
  try {
    var dropbox = "aaa";//แก้จุดที่ 11
    var folder, folders = DriveApp.getFoldersByName(dropbox);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(dropbox);
    }
    var contentType = data.substring(5,data.indexOf(';')),
        bytes = Utilities.base64Decode(data.substr(data.indexOf('base64,')+7)),
        blob = Utilities.newBlob(bytes, contentType, file),
        file = folder.createFile(blob);
    
        fileID = file.getId();
        Logger.log(fileID); 

        return ["OK",fileID];

  } catch (f) {
    return f.toString();
    return ContentService
  .createTextOutput(
    JSON.stringify({"result":"file upload failed",
                    "data": JSON.stringify(f) })) 
  .setMimeType(ContentService.MimeType.JSON);
  }
}


//////HTML////////
<!DOCTYPE html>
<html>
    <head>
        <base target="_top">
        <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" 
        integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">
           <link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Prompt">
<style>

body {
  font-family: "Prompt";
  font-size: 16;
}
</style>
        <?!= include('JavaScript'); ?>
    </head>
    <body>

        <div class="container  mt-3r">
            <div class="row" >
                <div class="col-6">
                    
                    
                        <img src="aaa" width="100%" class="img-fluid" alt="Responsive image"  >
                        <div class="alert alert-primary"  role="alert"><center>
                        <B><font size=4>ระบบการรับสมัครนักเรียน ประจำปีการศึกษา 2563</font></B>
                        </center></div>
                        <div class="alert alert-success" role="alert">
                        <h6 class="font-weight-bold">
                        คำแนะนำในการกรอกข้อมูลสมัครเข้าเรียน</h6>
                        <p>1. ให้นักเรียนกรอกข้อมูลที่เป็นจริงและถูกต้อง และตรวจสอบข้อมูลให้เรียบร้อยก่อนกดปุ่มบันทึก
                        <br>2. หากสงสัยหรือมีปัญหาในการสมัคร ให้สอบถามข้อมูลเพิ่มเติมได้ที่..โทร 089-9464558 </p>
                        </div>    
                        

<form  id="form" >
            <div class="form-row">
             <div class="form-group col-md-6">
              <label for="upload">รูปนักเรียน</label>
              <input id="files" type="file">
              <img id="image" />
            </div>
            <div class="form-group col-md-3">
            <label for="upload">คลิกอัปโหลดรูป</label>
                <button type="button" id="pic_name" name="pic_name" onclick="uploadImage(); return false;">Upload</button>
                </div>
                <div class="form-group col-md-3">
                <label for="upload"></label>
                <div id = "progress"></div>
    <div id="success" style="display:none">อัปโหลดเรียบร้อย..</div> </div></div>
    </form>  

        <form id="myForm" onsubmit="handleFormSubmit(this)">
<!--  ชื่อ สกุล   -->
                        <div class="form-row">
                            <div class="form-group col-md-6">
                                <label for="first_name">ชื่อ</label>
                                <input type="text" class="form-control" id="first_name" name="first_name" placeholder="ชื่อ">
                            </div>
                            <div class="form-group col-md-6">
                                <label for="last_name">สกุล</label>
                                <input type="text" class="form-control" id="last_name" name="last_name" placeholder="สกุล">
                            </div>
                        </div>
<!--  เลขประชาชน  -->

                        <div class="form-row">
                         <div class="form-group col-md-6">
                            <label for="iduser">เลขประชาชน</label>
                            <input type="text" class="form-control" maxlength="13" id="iduser" name="iduser" placeholder="เลขประชาชน">
                         </div>
                            <div class="form-group col-md-6">
                            <label for="email">อีเมล</label>
                            <input type="text" class="form-control"  id="email" name="email" placeholder="อีเมล">
                         </div>
                          </div>
                          
<!--  เพศ  -->
                        <div class="form-row">
                            <div class="form-group col-md-6">
                                <p>เพศ</p>
                                <div class="form-check form-check-inline">
                                    <input class="form-check-input" type="radio" name="gender" id="male" value="ชาย">
                                    <label class="form-check-label" for="male">ชาย</label>
                                </div>
                                <div class="form-check form-check-inline">
                                    <input class="form-check-input" type="radio" name="gender" id="female" value="หญิง">
                                    <label class="form-check-label" for="female">หญิง</label>
                                </div>
                            </div></div>
                            
 <!--  วันเกิด  -->  
                           <div class="form-row">
                            <div class="form-group col-md-6">
                                <label for="dateOfBirth">วันเกิด</label>
                                <input type="date" class="form-control" id="dateOfBirth" name="dateOfBirth">
                            </div>
                           
                        
 <!--  เบอร์โทร   -->
 
                        <div class="form-group col-md-6">
                            <label for="phone">เบอร์โทร</label>
                            <input type="tel" class="form-control" id="phone" maxlength="11" name="phone" placeholder="088-888xxxx">
                         </div>
                         </div>


 <!--  ที่อยู่   -->                         

                        <div class="form-row">
                        <label for="address">ที่อยู่</label>
                        <input type="text" class="form-control" id="address" name="address" placeholder="ที่อยู่">
                        </div>
                        
<!--  ชื่อเล่น   -->

                        <div class="form-row">
                         <div class="form-group col-md-6">
                            <label for="nickname">ชื่อเล่น</label>
                            <input type="text" class="form-control" id="nickname" name="nickname" placeholder="ชื่อเล่น">
                         </div>
                          </div>
                          
<!--  ระดับชั้น   -->                          
                          
                        <div class="form-row">
                            <div class="form-group ">
                                <p>ระดับชั้นที่สมัคร</p>
                                <div class="form-check form-check-inline">
                                    <input class="form-check-input" type="radio" name="class" id="class" value="มัธยมศึกษาปีที่ 1">
                                    <label class="form-check-label" for="class1">มัธยมศึกษาปีที่ 1</label>
                                </div>
                                <div class="form-check form-check-inline">
                                    <input class="form-check-input" type="radio" name="class" id="class" value="มัธยมศึกษาปีที่ 4">
                                    <label class="form-check-label" for="class4">มัธยมศึกษาปีที่ 4</label>
                                </div>
                                </div> 
                            </div>
                            
<!--  แผนการเรียน   -->                            
                            <div class="form-row">
                            <div class="form-group col-md-6">
                            <label for="planselect">แผนการเรียน</label>
                            <select id="planselect" class="form-control" name = "plan">
                            <option selected>เลือกแผนการเรียน</option>
                            <option>แผนการเรียนที่ 1</option>
                            <option>แผนการเรียนที่ 2</option>
                            <option>แผนการเรียนที่ 3</option>
                            </select>
                            </div> 
                             </div> 
<!--  กรุ๊บเลือด   --> 
                             <div class="form-row">
                            <div class="form-group col-md-6">
                            <label for="blood">กรุ๊บเลือด</label>
                            <select id="blood" class="form-control" name = "blood">
                            <option selected>เลือกกรุ๊บเลือด</option>
                            <option>กรุ๊ป A</option>
                            <option>กรุ๊ป B</option>
                            <option>กรุ๊ป AB</option>
                            <option>กรุ๊ป O</option>
                            </select>
                            </div> 
                             </div>  
                          
<!--  ปุ่มบันทึกข้อมูล   -->      
                     <input type="hidden" id="picID" name="picID">
                        <br>
                        <button type="submit" id = "save" onclick="if(confirm('ท่านยืนยันที่จะบันทึกข้อมูลหรือไม่ ?')){alert('บันทึกข้อมูลเรียบร้อยแล้ว!!')
}else{return false;};" class="btn btn-primary btn-block">บันทึกข้อมูล</button>
<br>
                    </form>
                     
                   <div class="alert alert-primary"  role="alert"><center>
                        <font size=3>จัดทำโดย</font><br>
                        <font size=3>นายอภิวัฒน์ วงศ์กัณหา ครู ชำนาญการพิเศษ</font>
                        </center></div>   
                    <div id="output"></div>
                </div>
            </div>      
        </div>
        
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/2.2.0/jquery.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/materialize/0.97.5/js/materialize.min.js"></script>
    <script>
         var file, 
          reader = new FileReader();
      reader.onloadend = function(e) {
        if (e.target.error != null) {
          showError("File " + file.name + " could not be read.");
          return;
        } else {
          google.script.run.withSuccessHandler(showSuccess)
          .uploadFileToGoogleDrive(e.target.result, file.name);
        }
      };

    document.getElementById("files").onchange = function () {
    var reader = new FileReader();
    reader.onload = function (e) {
    document.getElementById("image").src = e.target.result;
     $('#image').show();
     $('#image').size=100*100;
    };
    reader.readAsDataURL(this.files[0]);
    };

      function showSuccess(e) {
       if (e[0] === "OK") { 
          $('#progress').hide();
          $('#success').show();
          document.getElementById("save").disabled = false; 
          document.forms["myForm"]["picID"].value =e[1];
          document.getElementById("image").reset();
               } else {
          showError(e);
        }
      }


      function uploadImage() {
        var files = $('#files')[0].files;
        if (files.length === 0) {
          showError("กรุณาเลือกไฟล์");
          return;
        }
        file = files[0];
        if (file.size > 1024 * 1024 * 5) {
          showError("The file size should be < 5 MB.");
          return;
        }
        showMessage("กำลังอัปโหลด..");
        reader.readAsDataURL(file);
      }

      function showError(e) {
        $('#progress').addClass('red-text').html(e);
      }

      function showMessage(e) {
        $('#progress').removeClass('red-text').html(e);
      }
    </script>
    </body>
</html>



////////JAVASCRIPT//////
<script>
  // Prevent forms from submitting.
  function preventFormSubmit() {
    var forms = document.querySelectorAll('form');
    for (var i = 0; i < forms.length; i++) {
      forms[i].addEventListener('submit', function(event) {
      event.preventDefault();
      });
    }
  }
  window.addEventListener('load', preventFormSubmit);    
      
      
  function handleFormSubmit(formObject) {
    google.script.run.processForm(formObject);
    document.getElementById("myForm").reset();
    document.getElementById("form").reset(); 
    $('#progress').hide();  
    $('#success').hide();  
    $('#image').hide();
    document.getElementById("save").disabled = true; 

  }
</script>
