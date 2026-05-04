/**
 * ตั้งค่า ID ของ Google Sheets ที่ต้องการเชื่อมต่อ
 * (สามารถปล่อยว่างไว้ได้ หากโค้ดนี้ถูกสร้าง (Bound) อยู่ภายใน Google Sheets นั้นๆ อยู่แล้ว)
 */
const SPREADSHEET_ID = '1oxZgtI5p_1PJo2y6kdZgm3fuDj2hDl4sP9-eic8TGu0'; // ใส่ ID Sheet ที่นี่หากเป็น Standalone Script

// ฟังก์ชันเริ่มต้นสำหรับ Render หน้า Web App
function doGet(e) {
  let template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('ระบบโหวตรางวัลที่สุดประจำรุ่น')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Helper: ดึงออบเจ็กต์ Spreadsheet
function getSpreadsheet() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * โหลดข้อมูลเริ่มต้น (รายชื่อนักเรียนทั้งหมด) สำหรับหน้า Login และ Modal
 */
function getInitialData() {
  const ss = getSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  if(!userSheet) throw new Error("ไม่พบชีต Users");

  const data = userSheet.getDataRange().getValues();
  const headers = data[0];
  const users = [];

  // เริ่มจากแถว 2 (Index 1) เพื่อข้าม Header
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    if (row[0]) { // ตรวจสอบว่ามีรหัสประจำตัว
      users.push({
        id: String(row[0]),
        name: row[2],
        nickname: row[3],
        image: row[4] || ''
      });
    }
  }
  return users;
}

/**
 * ระบบ Login ตรวจสอบ ID และ PIN (รองรับการตั้งรหัสผ่านครั้งแรก)
 */
function login(id, pin) {
  const ss = getSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  const data = userSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      let savedPin = String(data[i][1] || "").trim(); // ดึงรหัสผ่านเดิมจากคอลัมน์ B

      // กรณีเข้าครั้งแรก (คอลัมน์ PIN ยังเป็นช่องว่าง)
      if (savedPin === "") {
        // บันทึกรหัสผ่านใหม่ที่ผู้ใช้เพิ่งพิมพ์ ลงไปในคอลัมน์ B (แถวที่ i+1, คอลัมน์ที่ 2)
        userSheet.getRange(i + 1, 2).setValue(pin);
        return {
          success: true,
          user: {
            id: String(data[i][0]),
            name: data[i][2],
            nickname: data[i][3],
            image: data[i][4] || ''
          }
        };
      }
      // กรณีเคยตั้งรหัสไว้แล้ว (ตรวจสอบว่าตรงกับที่บันทึกไว้ไหม)
      else if (savedPin === String(pin)) {
        return {
          success: true,
          user: {
            id: String(data[i][0]),
            name: data[i][2],
            nickname: data[i][3],
            image: data[i][4] || ''
          }
        };
      }
      // กรณีรหัสผ่านไม่ตรง
      else {
        return { success: false, message: 'รหัสผ่านไม่ถูกต้อง (หากจำรหัสไม่ได้ กรุณาติดต่อผู้ดูแล)' };
      }
    }
  }
  return { success: false, message: 'ไม่พบชื่อนี้ในระบบ' };
}

/**
 * ดึงข้อมูลการโหวตเดิมของผู้ใช้ (ถ้ามี)
 */
function getUserVotes(voterId) {
  const ss = getSpreadsheet();
  const voteSheet = ss.getSheetByName('Votes');
  if(!voteSheet) return {};

  const data = voteSheet.getDataRange().getValues();
  let userVotes = {};

  // ค้นหาจากล่างขึ้นบน เผื่อมีหลาย Record (ดึงล่าสุด)
  for (let i = data.length - 1; i > 0; i--) {
    if (String(data[i][1]) === String(voterId)) {
      // คอลัมน์ C ถึง Q (Index 2 ถึง 14) คือหมวดที่ 3 ถึง 15
      for (let j = 3; j <= 15; j++) {
        let voteValue = data[i][j - 1]; // Index หมวด 3 คือ 2, หมวด 4 คือ 3 ...
        if (voteValue) userVotes[`cat_${j}`] = String(voteValue);
      }
      break;
    }
  }
  return userVotes;
}

/**
 * บันทึกผลโหวต
 */
function submitVotes(voterId, votesObj) {
  const ss = getSpreadsheet();
  let voteSheet = ss.getSheetByName('Votes');
  
  // ถ้ายังไม่มีชีต Votes ให้สร้างใหม่พร้อม Header
  if (!voteSheet) {
    voteSheet = ss.insertSheet('Votes');
    voteSheet.appendRow(['Timestamp', 'VoterID', 'Cat3', 'Cat4', 'Cat5', 'Cat6', 'Cat7', 'Cat8', 'Cat9', 'Cat10', 'Cat11', 'Cat12', 'Cat13', 'Cat14', 'Cat15']);
  }

  const data = voteSheet.getDataRange().getValues();
  let rowIndexToUpdate = -1;

  // ค้นหาว่าเคยโหวตหรือยัง
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(voterId)) {
      rowIndexToUpdate = i + 1; // +1 เพราะ data array เริ่มที่ 0 แต่ row sheet เริ่มที่ 1
      break;
    }
  }

  // เตรียมข้อมูลแถวที่จะบันทึก
  const newRowData = [
    new Date(), // Timestamp
    voterId,    // VoterID
    votesObj.cat_3 || '', votesObj.cat_4 || '', votesObj.cat_5 || '',
    votesObj.cat_6 || '', votesObj.cat_7 || '', votesObj.cat_8 || '',
    votesObj.cat_9 || '', votesObj.cat_10 || '', votesObj.cat_11 || '',
    votesObj.cat_12 || '', votesObj.cat_13 || '', votesObj.cat_14 || '',
    votesObj.cat_15 || ''
  ];

  // Lock ป้องกันการบันทึกชนกัน
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (rowIndexToUpdate > -1) {
      // แก้ไขแถวเดิม
      voteSheet.getRange(rowIndexToUpdate, 1, 1, newRowData.length).setValues([newRowData]);
    } else {
      // เพิ่มแถวใหม่
      voteSheet.appendRow(newRowData);
    }
    return { success: true, message: 'บันทึกผลโหวตเรียบร้อยแล้ว' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * คำนวณผลคะแนน Leaderboard (Top 3 ของแต่ละหมวด) แบบ Anonymous
 */
function getLeaderboard() {
  const ss = getSpreadsheet();
  const voteSheet = ss.getSheetByName('Votes');
  const userSheet = ss.getSheetByName('Users');
  
  if (!voteSheet || !userSheet) return {};

  const votesData = voteSheet.getDataRange().getValues();
  const usersData = userSheet.getDataRange().getValues();
  
  // สร้าง User Dictionary สำหรับดึงข้อมูลรูปและชื่ออย่างรวดเร็ว
  const userDict = {};
  for(let i = 1; i < usersData.length; i++){
    userDict[String(usersData[i][0])] = {
      nickname: usersData[i][3],
      image: usersData[i][4]
    };
  }

  // นับคะแนน
  const scoreMap = {}; // { cat_3: { userId1: 5, userId2: 2 }, cat_4: ... }
  for(let i = 3; i <= 15; i++) {
    scoreMap[`cat_${i}`] = {};
  }

  for (let i = 1; i < votesData.length; i++) {
    for (let j = 3; j <= 15; j++) {
      let votedFor = String(votesData[i][j - 1]);
      if (votedFor && votedFor !== '') {
        let catKey = `cat_${j}`;
        scoreMap[catKey][votedFor] = (scoreMap[catKey][votedFor] || 0) + 1;
      }
    }
  }

  // จัดเรียงและหา Top 3
  const leaderboard = {};
  for(let i = 3; i <= 15; i++) {
    let catKey = `cat_${i}`;
    let catScores = scoreMap[catKey];
    
    // แปลงเป็น Array แล้วเรียงจากมากไปน้อย
    let sortedScores = Object.keys(catScores).map(id => {
      return {
        id: id,
        count: catScores[id],
        nickname: userDict[id] ? userDict[id].nickname : 'Unknown',
        image: userDict[id] ? userDict[id].image : ''
      };
    }).sort((a, b) => b.count - a.count).slice(0, 3); // เอาแค่ Top 3

    leaderboard[catKey] = sortedScores;
  }

  return leaderboard;
}
