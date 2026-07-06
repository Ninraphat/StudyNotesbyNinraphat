/* =====================================================================
   ระบบบอท LINE OA สำหรับร้านขายชีท — กึ่งอัตโนมัติ
   -----------------------------------------------------------------
   วิธีติดตั้ง (ทำครั้งเดียว):

   1) เปิด Google Sheet ใหม่ 1 ไฟล์ ตั้งชื่อ "Orders" ที่แถวแรกใส่หัวตาราง:
      OrderID | Timestamp | Items | Total | LineUserId | SlipReceived | Status | SentAt

   2) เข้า Extensions > Apps Script วางโค้ดนี้ทั้งหมดแทนของเดิม

   3) ไปที่ LINE Official Account Manager ของคุณ > ตั้งค่า > Messaging API
      - เปิดใช้งาน Messaging API (ถ้ายังไม่เปิด)
      - คัดลอก "Channel access token" (กด Issue ถ้ายังไม่มี) มาใส่ตัวแปร CHANNEL_ACCESS_TOKEN ด้านล่าง
      - ปิด "Auto-reply messages" และ "Greeting messages" ของ OA (กันชนกับบอท)

   4) ใน Apps Script กด Deploy > New deployment > เลือกประเภท "Web app"
      - Execute as: Me
      - Who has access: Anyone
      - กด Deploy จะได้ URL แบบ https://script.google.com/macros/s/XXXX/exec
      - เอา URL นี้ไปใส่ทั้งใน:
          a) index.html -> CONFIG.gasWebAppUrl
          b) LINE Developers Console (https://developers.line.biz) > channel ของ OA คุณ
             > Messaging API tab > Webhook URL > ใส่ URL เดียวกัน แล้วกด Verify + เปิด "Use webhook"

   5) หาเลข LINE userId ของตัวเอง (แอดมิน) เพื่อให้บอทรู้ว่าใครสั่งอนุมัติได้:
      - แอดเพื่อน OA ตัวเอง แล้วพิมพ์อะไรก็ได้ไปหาบอท 1 ครั้ง (ตอนนี้ยังไม่ทำอะไรก็ไม่เป็นไร)
      - เปิด Sheet ชื่อ "DebugLog" (โค้ดจะสร้างให้อัตโนมัติ) จะเห็น userId ของคุณ
      - คัดลอกมาใส่ใน ADMIN_USER_IDS ด้านล่าง แล้ว Deploy ใหม่อีกครั้ง (Manage deployments > Edit > New version)

   คำสั่งที่แอดมินใช้ในแชท LINE (พิมพ์คุยกับ OA ตัวเองในฐานะแอดมิน):
      /ส่ง ORDER25070512345     -> ตรวจสอบสลิปแล้ว ให้บอทส่งลิงก์ไฟล์ไปหาลูกค้าอัตโนมัติ
      /ดู ORDER25070512345      -> ดูรายละเอียดออเดอร์นั้น
===================================================================== */

const CHANNEL_ACCESS_TOKEN = "YOUR_CHANNEL_ACCESS_TOKEN"; // จาก LINE Developers Console
const ADMIN_USER_IDS = ["YOUR_LINE_USER_ID"]; // userId ของแอดมิน (คุณ) เท่านั้นที่สั่งอนุมัติได้
const SHEET_ID = "YOUR_GOOGLE_SHEET_ID"; // เปิด Sheet แล้วก็อปจาก URL ช่วง /d/XXXX/edit

// ต้องตรงกับ CONFIG.subjects ในเว็บ index.html
const SUBJECT_LINKS = {
  "ENG2001": { link:"https://drive.google.com/FILE_ID_ENG2001" },
  "ENG2002": { link:"https://drive.google.com/FILE_ID_ENG2002", addonLink:"https://drive.google.com/FILE_ID_ENG2002_FLASHCARD", addonLabel:"Flashcard Quizlet" },
  "ENG2101": { link:"https://drive.google.com/FILE_ID_ENG2101", addonLink:"https://drive.google.com/FILE_ID_ENG2101_EXAM", addonLabel:"แนวข้อสอบเพิ่มเติม" },
  "ENG2102": { link:"https://drive.google.com/FILE_ID_ENG2102" },
  "ENL2001": { link:"https://drive.google.com/FILE_ID_ENL2001" },
  "ENG2401": { link:"https://drive.google.com/FILE_ID_ENG2401", addonLink:"https://drive.google.com/FILE_ID_ENG2401_EXAM", addonLabel:"แนวข้อสอบ 3 ชุด" },
  "ENG2601": { link:"https://drive.google.com/FILE_ID_ENG2601" },
};

function getSheet(name){
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if(!sheet){
    sheet = ss.insertSheet(name);
    if(name === "Orders"){
      sheet.appendRow(["OrderID","Timestamp","Items","Total","LineUserId","SlipReceived","Status","SentAt"]);
    }
  }
  return sheet;
}

function debugLog(msg){
  const sheet = getSheet("DebugLog");
  sheet.appendRow([new Date(), msg]);
}

function findOrderRow(sheet, orderId){
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){
    if(data[i][0] === orderId) return i+1; // แถวจริงใน Sheet (1-indexed)
  }
  return -1;
}

/* -------------------- รับ request จากเว็บ (สร้างออเดอร์) -------------------- */
function doPost(e){
  try{
    const body = JSON.parse(e.postData.contents);

    if(body.action === "create_order"){
      const sheet = getSheet("Orders");
      sheet.appendRow([body.orderId, new Date(), body.items, body.total, "", false, "รอชำระเงิน", ""]);
      return ContentService.createTextOutput(JSON.stringify({ok:true}));
    }

    // ไม่ใช่ create_order แปลว่าเป็น webhook จาก LINE
    if(body.events){
      body.events.forEach(handleLineEvent);
      return ContentService.createTextOutput(JSON.stringify({ok:true}));
    }

    return ContentService.createTextOutput(JSON.stringify({ok:false, error:"unknown action"}));
  }catch(err){
    debugLog("doPost error: " + err);
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)}));
  }
}

/* -------------------- จัดการ event จาก LINE -------------------- */
function handleLineEvent(event){
  const userId = event.source && event.source.userId;
  if(!userId) return;
  debugLog("event from userId: " + userId + " | type: " + event.type);

  if(event.type !== "message") return;

  const isAdmin = ADMIN_USER_IDS.indexOf(userId) !== -1;

  // ---------- ข้อความจากแอดมิน: คำสั่ง /ส่ง หรือ /ดู ----------
  if(isAdmin && event.message.type === "text"){
    const text = event.message.text.trim();

    if(text.startsWith("/ส่ง")){
      const orderId = text.replace("/ส่ง","").trim();
      approveAndSendLinks(orderId, event.replyToken);
      return;
    }
    if(text.startsWith("/ดู")){
      const orderId = text.replace("/ดู","").trim();
      showOrderDetail(orderId, event.replyToken);
      return;
    }
  }

  // ---------- ข้อความจากลูกค้า ----------
  if(event.message.type === "text"){
    const text = event.message.text;
    const match = text.match(/ORDER[0-9]+/);
    if(match){
      const orderId = match[0];
      const sheet = getSheet("Orders");
      const row = findOrderRow(sheet, orderId);
      if(row > 0){
        sheet.getRange(row, 5).setValue(userId); // บันทึก LineUserId คู่กับออเดอร์นี้
        replyMessage(event.replyToken, `รับออเดอร์ ${orderId} แล้วค่ะ ✅\nกรุณาแนบรูปสลิปการโอนในแชทนี้ได้เลย รอแอดมินตรวจสอบสักครู่นะคะ`);
      } else {
        replyMessage(event.replyToken, `ไม่พบเลขออเดอร์ ${orderId} ในระบบ กรุณาตรวจสอบอีกครั้งค่ะ`);
      }
    }
  }

  if(event.message.type === "image"){
    // มาร์คว่าลูกค้าคนนี้ส่งสลิปแล้ว (จับคู่จาก LineUserId ล่าสุดที่ตรงกับออเดอร์ที่ยังไม่ส่ง)
    const sheet = getSheet("Orders");
    const data = sheet.getDataRange().getValues();
    for(let i=data.length-1; i>=1; i--){
      if(data[i][4] === userId && data[i][6] !== "ส่งแล้ว"){
        sheet.getRange(i+1, 6).setValue(true);
        sheet.getRange(i+1, 7).setValue("รอตรวจสอบสลิป");
        replyMessage(event.replyToken, `ได้รับสลิปแล้วค่ะ 🧾 รอแอดมินตรวจสอบและส่งลิงก์ให้นะคะ (ปกติไม่เกิน 1-2 ชม.)`);
        notifyAdmins(`📥 มีสลิปใหม่\nออเดอร์: ${data[i][0]}\nยอด: ${data[i][3]} บาท\nพิมพ์ /ส่ง ${data[i][0]} เพื่ออนุมัติ`);
        return;
      }
    }
  }
}

/* -------------------- แอดมินอนุมัติ -> ส่งลิงก์ให้ลูกค้า -------------------- */
function approveAndSendLinks(orderId, replyToken){
  const sheet = getSheet("Orders");
  const row = findOrderRow(sheet, orderId);
  if(row < 0){
    replyMessage(replyToken, `ไม่พบออเดอร์ ${orderId}`);
    return;
  }
  const rowData = sheet.getRange(row, 1, 1, 8).getValues()[0];
  const [ , , items, total, lineUserId, , status] = rowData;

  if(!lineUserId){
    replyMessage(replyToken, `ออเดอร์ ${orderId} ยังไม่มี LineUserId ของลูกค้า (ลูกค้ายังไม่ได้พิมพ์เลขออเดอร์ในแชท)`);
    return;
  }
  if(status === "ส่งแล้ว"){
    replyMessage(replyToken, `ออเดอร์ ${orderId} ส่งลิงก์ไปแล้วก่อนหน้านี้`);
    return;
  }

  const codes = items.split(",").map(s => s.trim());
  let msg = `ชำระเงินสำเร็จค่ะ 🎉\nนี่คือลิงก์ไฟล์ของคุณ:\n\n`;
  codes.forEach(code => {
    const hasAddon = code.includes("+ADDON");
    const cleanCode = code.replace("+ADDON","").trim();
    const s = SUBJECT_LINKS[cleanCode];
    if(!s) return;
    msg += `📄 ${cleanCode}: ${s.link}\n`;
    if(hasAddon && s.addonLink){
      msg += `📝 ${cleanCode} (${s.addonLabel}): ${s.addonLink}\n`;
    }
  });
  msg += `\nขอบคุณที่อุดหนุนค่ะ 🙏`;

  pushMessage(lineUserId, msg);
  sheet.getRange(row, 7).setValue("ส่งแล้ว");
  sheet.getRange(row, 8).setValue(new Date());

  replyMessage(replyToken, `ส่งลิงก์ให้ออเดอร์ ${orderId} เรียบร้อยแล้ว ✅`);
}

function showOrderDetail(orderId, replyToken){
  const sheet = getSheet("Orders");
  const row = findOrderRow(sheet, orderId);
  if(row < 0){
    replyMessage(replyToken, `ไม่พบออเดอร์ ${orderId}`);
    return;
  }
  const rowData = sheet.getRange(row, 1, 1, 8).getValues()[0];
  const [id, timestamp, items, total, lineUserId, slipReceived, status] = rowData;
  const msg = `ออเดอร์: ${id}\nเวลา: ${timestamp}\nรายการ: ${items}\nยอด: ${total} บาท\nได้รับสลิป: ${slipReceived ? "ใช่" : "ยังไม่ได้รับ"}\nสถานะ: ${status}`;
  replyMessage(replyToken, msg);
}

/* -------------------- LINE API helpers -------------------- */
function replyMessage(replyToken, text){
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
    method: "post",
    headers: { "Authorization": "Bearer " + CHANNEL_ACCESS_TOKEN, "Content-Type": "application/json" },
    payload: JSON.stringify({ replyToken, messages: [{ type:"text", text }] }),
    muteHttpExceptions: true
  });
}

function pushMessage(userId, text){
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
    method: "post",
    headers: { "Authorization": "Bearer " + CHANNEL_ACCESS_TOKEN, "Content-Type": "application/json" },
    payload: JSON.stringify({ to: userId, messages: [{ type:"text", text }] }),
    muteHttpExceptions: true
  });
}

function notifyAdmins(text){
  ADMIN_USER_IDS.forEach(id => pushMessage(id, text));
}
