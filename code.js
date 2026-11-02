// 1. ฟังก์ชั่นสร้าง/เตรียมหัวตาราง Google Sheet อัตโนมัติ (รันแค่นี้ครั้งแรกครั้งเดียว)
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // สร้าง Sheet1: Posts
  var postSheet = ss.getSheetByName("Posts") || ss.insertSheet("Posts");
  if (postSheet.getLastRow() === 0) {
    postSheet.appendRow(["PostID", "Timestamp", "MemberID", "Platform", "ReelUrl", "Title", "TargetLikes"]);
  }
  
  // สร้าง Sheet2: InteractionLogs
  var logSheet = ss.getSheetByName("InteractionLogs") || ss.insertSheet("InteractionLogs");
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(["PostID", "MemberID", "Timestamp", "Status"]);
  }
  
  Logger.log("Setup completed successfully!");
}

// 2. รับการเรียก API แบบ GET (สำหรับดึงข้อมูลไปแสดงหน้าเว็บ)
function doGet(e) {
  var action = e.parameter.action;
  
  if (action === "getDashboard") {
    return respondJSON(getDashboardData());
  }
  
  return respondJSON({ status: "error", message: "Invalid action" });
}

// 3. รับการเรียก API แบบ POST (สำหรับบันทึกโพสต์ หรือ บันทึกการกดไลค์)
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    
    if (action === "addPost") {
      return respondJSON(addPost(data));
    } else if (action === "toggleLike") {
      return respondJSON(toggleLike(data));
    }
    
    return respondJSON({ status: "error", message: "Unknown action" });
  } catch (err) {
    return respondJSON({ status: "error", message: err.toString() });
  }
}

// Helper: บันทึกโพสต์ใหม่
function addPost(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Posts");
  var postId = "P" + new Date().getTime();
  
  sheet.appendRow([
    postId,
    new Date().toISOString(),
    data.memberId,
    data.platform,
    data.reelUrl,
    data.title || "Reel คลิปใหม่",
    data.targetLikes || 50
  ]);
  
  return { status: "success", postId: postId };
}

// Helper: บันทึกการกดไลค์ของสมาชิก
function toggleLike(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("InteractionLogs");
  var rows = sheet.getDataRange().getValues();
  
  // ตรวจสอบว่าเคยบันทึกไว้หรือยัง
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.postId && rows[i][1] === data.memberId) {
      // ถ้าเคยมีแล้ว ลบออก (Toggle Off)
      sheet.deleteRow(i + 1);
      return { status: "success", liked: false };
    }
  }
  
  // ถ้ายังไม่มี ให้บันทึกเพิ่ม (Toggle On)
  sheet.appendRow([data.postId, data.memberId, new Date().toISOString(), "LIKED"]);
  return { status: "success", liked: true };
}

// Helper: ดึงข้อมูลสรุปไปแสดงผล
function getDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var postsSheet = ss.getSheetByName("Posts");
  var logsSheet = ss.getSheetByName("InteractionLogs");
  
  var postsData = postsSheet.getDataRange().getValues();
  var logsData = logsSheet.getDataRange().getValues();
  
  var posts = [];
  
  // วนลูปอ่านโพสต์ทั้งหมด
  for (var i = 1; i < postsData.length; i++) {
    var postId = postsData[i][0];
    var likedMembers = [];
    
    // ค้นหาสมาชิกที่กดไลค์คลิปนี้แล้ว
    for (var j = 1; j < logsData.length; j++) {
      if (logsData[j][0] === postId) {
        likedMembers.push(logsData[j][1]);
      }
    }
    
    posts.push({
      postId: postId,
      timestamp: postsData[i][1],
      memberId: postsData[i][2],
      platform: postsData[i][3],
      reelUrl: postsData[i][4],
      title: postsData[i][5],
      targetLikes: postsData[i][6],
      currentLikes: likedMembers.length,
      likedMembers: likedMembers
    });
  }
  
  return { status: "success", posts: posts.reverse() }; // เอาโพสต์ล่าสุดขึ้นก่อน
}

function respondJSON(object) {
  return ContentService.createTextOutput(JSON.stringify(object))
    .setMimeType(ContentService.MimeType.JSON);
}