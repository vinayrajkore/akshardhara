// ============================================================
// AKSHARDHARA VACHNALAY - Library Management System
// Google Apps Script Backend (Code.gs)
// ============================================================
// SETUP STEPS:
//   1. Open script.google.com → New Project → paste this file
//   2. Script Properties → Add:
//        BREVO_API_KEY       = your Brevo API key
//        BREVO_SENDER_EMAIL  = your verified sender email
//        BREVO_SENDER_NAME   = Akshardhara Vachnalay
//   3. Run setup() once to create all sheets
//   4. Run installDailyTrigger() for daily 9AM reminders
//   5. Deploy as Web App: Execute as Me, Access: Anyone
// ============================================================

var SHEET = {
  USERS: 'Users', OTP: 'OTPStore', BOOKS: 'Books',
  BORROWED: 'BorrowedBooks', REQUESTS: 'BookRequests', ADMIN: 'AdminMobiles'
};

// ---- Users sheet columns ----
// [0]ID [1]Name [2]Email [3]Mobile [4]Age [5]AuthProvider
// [6]GoogleID [7]Password(SHA256) [8]IsVerified [9]Role [10]CreatedDate

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configs = [
    { name: SHEET.USERS,    headers: ['ID','Name','Email','Mobile','Age','AuthProvider','GoogleID','Password','IsVerified','Role','CreatedDate'] },
    { name: SHEET.OTP,      headers: ['ID','Email','OTP','CreatedAt','ExpiresAt','IsUsed'] },
    { name: SHEET.BOOKS,    headers: ['ID','BookName','AuthorName','Publication','Price','IsAvailable','AddedDate'] },
    { name: SHEET.BORROWED, headers: ['ID','BookID','BookName','ReaderName','MobileNumber','BorrowDate','DueDate','ReturnDate','Status','UserID'] },
    { name: SHEET.REQUESTS, headers: ['ID','BookName','AuthorName','RequestedBy','MobileNumber','RequestDate','Status','UserID'] },
    { name: SHEET.ADMIN,    headers: ['MobileNumber','Name','AddedDate'] }
  ];
  configs.forEach(function(c) {
    var s = ss.getSheetByName(c.name) || ss.insertSheet(c.name);
    if (s.getLastRow() === 0) {
      s.appendRow(c.headers);
      s.getRange(1,1,1,c.headers.length).setFontWeight('bold').setBackground('#7c1d1d').setFontColor('#fef3c7');
      s.setFrozenRows(1);
    }
  });
  var def = ss.getSheetByName('Sheet1');
  if (def) ss.deleteSheet(def);
  return { success: true, message: 'Setup complete! All 6 sheets created.' };
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var a = e.parameter.action, r;
    switch(a) {
      case 'getBooks':      r = getBooks(); break;
      case 'getAllBorrowed': r = getAllBorrowedBooks(); break;
      case 'getMyBooks':    r = getMyBooks(e.parameter.userID); break;
      case 'getRequests':   r = getRequests(); break;
      case 'getDashboard':  r = getDashboardStats(e.parameter.userID); break;
      case 'getAllUsers':   r = getAllUsers(e.parameter.userID); break;
      case 'ping':          r = { success: true, message: 'OK' }; break;
      default:              r = { error: 'Unknown action' };
    }
    return respond(r);
  } catch(err) { return respond({ error: err.toString() }); }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents), r;
    switch(data.action) {
      case 'googleAuth':         r = googleAuth(data); break;
      case 'sendOTP':            r = sendOTP(data.email); break;
      case 'verifyOTP':          r = verifyOTP(data.email, data.otp); break;
      case 'registerEmail':      r = registerEmail(data); break;
      case 'loginEmail':         r = loginEmail(data); break;
      case 'borrowBook':         r = borrowBook(data); break;
      case 'returnBook':         r = returnBook(data); break;
      case 'requestBook':        r = requestBook(data); break;
      case 'addBook':            r = addBook(data); break;
      case 'updateBook':         r = updateBook(data); break;
      case 'deleteBook':         r = deleteBook(data); break;
      case 'fulfillRequest':     r = fulfillRequest(data); break;
      case 'rejectRequest':      r = rejectRequest(data); break;
      case 'updateUserRole':     r = updateUserRole(data); break;
      case 'sendManualReminder': r = sendManualReminder(data); break;
      default:                   r = { error: 'Unknown action' };
    }
    return respond(r);
  } catch(err) { return respond({ error: err.toString() }); }
}

// ============================================================
// AUTH
// ============================================================

function googleAuth(data) {
  var email = data.email, googleID = data.googleID, name = data.name;
  if (!email || !googleID) return { error: 'Missing Google credentials' };
  var sheet = getSheet(SHEET.USERS), rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][2] === email) return { success: true, user: rowToUser(rows[i]), isNew: false };
  }
  var id = generateID(), now = new Date().toISOString();
  sheet.appendRow([id, name, email, '', '', 'google', googleID, '', true, 'reader', now]);
  return { success: true, user: { id:id, name:name, email:email, mobile:'', age:'', role:'reader', authProvider:'google' }, isNew: true };
}

function sendOTP(email) {
  if (!email) return { error: 'Email is required' };
  var otp = generateOTP(), now = new Date(), expires = new Date(now.getTime() + 600000);
  var sheet = getSheet(SHEET.OTP), rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][1] === email && rows[i][5] === false) sheet.getRange(i+1,6).setValue(true);
  }
  sheet.appendRow([generateID(), email, otp, now.toISOString(), expires.toISOString(), false]);
  var html = '<div style="font-family:Georgia,serif;max-width:520px;margin:auto;background:#0e0a06;padding:40px;border-radius:16px;border:1px solid rgba(200,164,74,0.3);">'
    + '<div style="text-align:center;margin-bottom:28px;">'
    + '<div style="width:70px;height:70px;background:linear-gradient(135deg,#c8a44a,#8b6914);border-radius:50%;margin:0 auto 16px;display:flex;align-items:center;justify-content:center;font-size:32px;">📚</div>'
    + '<h2 style="color:#e8c97a;margin:0;font-size:22px;letter-spacing:1px;">अक्षरधारा वाचनालय</h2>'
    + '<p style="color:#b8a88a;margin:4px 0 0;font-size:13px;">Email Verification</p></div>'
    + '<p style="color:#c9b99a;font-size:15px;">Your One-Time Verification Code:</p>'
    + '<div style="background:linear-gradient(135deg,rgba(200,164,74,0.15),rgba(139,30,30,0.1));border:1px solid rgba(200,164,74,0.4);border-radius:12px;padding:28px;text-align:center;margin:20px 0;">'
    + '<span style="color:#e8c97a;font-size:48px;font-weight:800;letter-spacing:14px;font-family:monospace;">' + otp + '</span></div>'
    + '<p style="color:#7a6b54;font-size:13px;">⏱ Valid for <strong style="color:#c8a44a;">10 minutes</strong>. Please do not share this code.</p>'
    + '</div>';
  var r = sendBrevoEmail(email, 'Akshardhara Vachnalay – Email Verification', html);
  if (!r.success) return { error: 'Failed to send OTP: ' + r.error };
  return { success: true, message: 'OTP sent to ' + email };
}

function verifyOTP(email, otp) {
  if (!email || !otp) return { error: 'Email and OTP required' };
  var sheet = getSheet(SHEET.OTP), rows = sheet.getDataRange().getValues(), now = new Date();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (rows[i][1] === email && String(rows[i][2]) === String(otp) && rows[i][5] === false) {
      if (now > new Date(rows[i][4])) return { error: 'OTP expired. Please request a new one.' };
      sheet.getRange(i+1,6).setValue(true);
      return { success: true, verified: true };
    }
  }
  return { error: 'Invalid OTP. Please try again.' };
}

// Register: OTP verified → create account with password
function registerEmail(data) {
  var name = data.name, email = data.email, mobile = data.mobile,
      age = data.age || '', password = data.password;
  if (!name || !email || !mobile || !password)
    return { error: 'Name, email, mobile, and password are required' };
  var sheet = getSheet(SHEET.USERS), rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][2] === email) return { error: 'Email already registered. Please login.' };
  }
  var id = generateID(), now = new Date().toISOString();
  var hashedPw = hashPassword(password);
  sheet.appendRow([id, name, email, mobile, age, 'email', '', hashedPw, true, 'reader', now]);
  return { success: true, user: { id:id, name:name, email:email, mobile:mobile, age:age, role:'reader', authProvider:'email' }, isNew: true };
}

// Login: email + password only (no OTP)
function loginEmail(data) {
  var email = data.email, password = data.password;
  if (!email || !password) return { error: 'Email and password are required' };
  var sheet = getSheet(SHEET.USERS), rows = sheet.getDataRange().getValues();
  var hashedPw = hashPassword(password);
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][2] === email) {
      if (rows[i][5] === 'google')
        return { error: 'This account uses Google Sign-In. Please use "Continue with Google" button.' };
      if (rows[i][7] !== hashedPw)
        return { error: 'Incorrect password. Please try again.' };
      return { success: true, user: rowToUser(rows[i]), isNew: false };
    }
  }
  return { success: false, needsRegistration: true };
}

// ============================================================
// LIBRARY
// ============================================================

function getBooks() {
  var rows = getSheet(SHEET.BOOKS).getDataRange().getValues(), books = [];
  for (var i = 1; i < rows.length; i++)
    books.push({ id:rows[i][0], bookName:rows[i][1], authorName:rows[i][2], publication:rows[i][3], price:rows[i][4], isAvailable:rows[i][5], addedDate:rows[i][6] });
  return { success: true, books: books };
}

function getAllBorrowedBooks() {
  var rows = getSheet(SHEET.BORROWED).getDataRange().getValues(), list = [];
  for (var i = 1; i < rows.length; i++)
    list.push({ id:rows[i][0], bookID:rows[i][1], bookName:rows[i][2], readerName:rows[i][3], mobile:rows[i][4], borrowDate:rows[i][5], dueDate:rows[i][6], returnDate:rows[i][7], status:rows[i][8], userID:rows[i][9] });
  return { success: true, borrowed: list };
}

function getMyBooks(userID) {
  if (!userID) return { error: 'UserID required' };
  var rows = getSheet(SHEET.BORROWED).getDataRange().getValues(), list = [];
  for (var i = 1; i < rows.length; i++)
    if (rows[i][9] === userID)
      list.push({ id:rows[i][0], bookID:rows[i][1], bookName:rows[i][2], borrowDate:rows[i][5], dueDate:rows[i][6], returnDate:rows[i][7], status:rows[i][8] });
  return { success: true, myBooks: list };
}

function borrowBook(data) {
  if (!data.bookID || !data.readerName || !data.mobile || !data.userID)
    return { error: 'All fields required' };
  var bSheet = getSheet(SHEET.BOOKS), books = bSheet.getDataRange().getValues(), bRow = -1, bName = '';
  for (var i = 1; i < books.length; i++) {
    if (books[i][0] === data.bookID) {
      if (!books[i][4]) return { error: 'Book is currently not available' };
      bRow = i+1; bName = books[i][1]; break;
    }
  }
  if (bRow === -1) return { error: 'Book not found' };
  bSheet.getRange(bRow, 5).setValue(false);
  var now = new Date(), due = new Date(now.getTime() + 10*86400000), id = generateID();
  getSheet(SHEET.BORROWED).appendRow([id, data.bookID, bName, data.readerName, data.mobile, now.toISOString(), due.toISOString(), '', 'borrowed', data.userID]);
  return { success: true, message: 'Book borrowed successfully!', borrowID: id, dueDate: due.toISOString() };
}

function returnBook(data) {
  if (!data.borrowID || !data.bookID) return { error: 'BorrowID and BookID required' };
  if (!isAdmin(data.requesterUserID)) return { error: 'Only admins can mark books as returned' };
  var bSheet = getSheet(SHEET.BORROWED), rows = bSheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.borrowID) { bSheet.getRange(i+1,8).setValue(new Date().toISOString()); bSheet.getRange(i+1,9).setValue('returned'); break; }
  }
  var books = getSheet(SHEET.BOOKS).getDataRange().getValues(), kSheet = getSheet(SHEET.BOOKS);
  for (var j = 1; j < books.length; j++) {
    if (books[j][0] === data.bookID) { kSheet.getRange(j+1,5).setValue(true); break; }
  }
  return { success: true, message: 'Book returned successfully!' };
}

function requestBook(data) {
  if (!data.bookName || !data.requestedBy || !data.userID) return { error: 'Book name and login required' };
  var id = generateID();
  getSheet(SHEET.REQUESTS).appendRow([id, data.bookName, data.authorName||'', data.requestedBy, data.mobile||'', new Date().toISOString(), 'pending', data.userID]);
  return { success: true, message: 'Book request submitted!' };
}

function getRequests() {
  var rows = getSheet(SHEET.REQUESTS).getDataRange().getValues(), list = [];
  for (var i = 1; i < rows.length; i++)
    list.push({ id:rows[i][0], bookName:rows[i][1], authorName:rows[i][2], requestedBy:rows[i][3], mobile:rows[i][4], requestDate:rows[i][5], status:rows[i][6], userID:rows[i][7] });
  return { success: true, requests: list };
}

// ============================================================
// ADMIN
// ============================================================

function addBook(data) {
  if (!data.bookName || !data.authorName) return { error: 'Book name and author required' };
  if (!isAdmin(data.requesterUserID)) return { error: 'Unauthorized' };
  var id = generateID();
  getSheet(SHEET.BOOKS).appendRow([id, data.bookName, data.authorName, data.publication||'', data.price||0, true, new Date().toISOString()]);
  return { success: true, message: 'Book added!', id: id };
}

function updateBook(data) {
  if (!data.bookID) return { error: 'BookID required' };
  if (!isAdmin(data.requesterUserID)) return { error: 'Unauthorized' };
  var sheet = getSheet(SHEET.BOOKS), rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.bookID) {
      if (data.bookName !== undefined)    sheet.getRange(i+1,2).setValue(data.bookName);
      if (data.authorName !== undefined)  sheet.getRange(i+1,3).setValue(data.authorName);
      if (data.publication !== undefined) sheet.getRange(i+1,4).setValue(data.publication);
      if (data.price !== undefined)       sheet.getRange(i+1,5).setValue(data.price);
      if (data.isAvailable !== undefined) sheet.getRange(i+1,6).setValue(data.isAvailable);
      return { success: true, message: 'Book updated!' };
    }
  }
  return { error: 'Book not found' };
}

function deleteBook(data) {
  if (!data.bookID) return { error: 'BookID required' };
  if (!isAdmin(data.requesterUserID)) return { error: 'Unauthorized' };
  var sheet = getSheet(SHEET.BOOKS), rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.bookID) { sheet.deleteRow(i+1); return { success: true, message: 'Book deleted!' }; }
  }
  return { error: 'Book not found' };
}

function fulfillRequest(data) {
  if (!isAdmin(data.requesterUserID)) return { error: 'Unauthorized' };
  var sheet = getSheet(SHEET.REQUESTS), rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.requestID) { sheet.getRange(i+1,7).setValue('fulfilled'); return { success: true, message: 'Request fulfilled!' }; }
  }
  return { error: 'Request not found' };
}

function rejectRequest(data) {
  if (!isAdmin(data.requesterUserID)) return { error: 'Unauthorized' };
  var sheet = getSheet(SHEET.REQUESTS), rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.requestID) { sheet.getRange(i+1,7).setValue('rejected'); return { success: true, message: 'Request rejected.' }; }
  }
  return { error: 'Request not found' };
}

function getDashboardStats(requesterUserID) {
  if (!isAdmin(requesterUserID)) return { error: 'Unauthorized' };
  var books = getSheet(SHEET.BOOKS).getDataRange().getValues();
  var borrowed = getSheet(SHEET.BORROWED).getDataRange().getValues();
  var requests = getSheet(SHEET.REQUESTS).getDataRange().getValues();
  var users = getSheet(SHEET.USERS).getDataRange().getValues();
  var now = new Date();
  var overdue = borrowed.slice(1).filter(function(b){ return b[8]==='borrowed' && now > new Date(b[6]); })
    .map(function(b){ return { id:b[0], bookName:b[2], readerName:b[3], mobile:b[4], dueDate:b[6], userID:b[9] }; });
  return {
    success: true,
    stats: { totalBooks: Math.max(0,books.length-1), availableBooks: books.slice(1).filter(function(b){return b[5]===true;}).length, activeBorrows: borrowed.slice(1).filter(function(b){return b[8]==='borrowed';}).length, pendingRequests: requests.slice(1).filter(function(r){return r[6]==='pending';}).length, totalUsers: Math.max(0,users.length-1) },
    overdue: overdue
  };
}

function getAllUsers(requesterUserID) {
  if (!isAdmin(requesterUserID)) return { error: 'Unauthorized' };
  var rows = getSheet(SHEET.USERS).getDataRange().getValues(), list = [];
  for (var i = 1; i < rows.length; i++) list.push(rowToUser(rows[i]));
  return { success: true, users: list };
}

function updateUserRole(data) {
  if (!isAdmin(data.requesterUserID)) return { error: 'Unauthorized' };
  var sheet = getSheet(SHEET.USERS), rows = sheet.getDataRange().getValues();
  if (data.newRole === 'admin') {
    var cnt = rows.slice(1).filter(function(r){ return r[9]==='admin' && r[0]!==data.targetUserID; }).length;
    if (cnt >= 2) return { error: 'Maximum 2 admin accounts allowed' };
  }
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.targetUserID) { sheet.getRange(i+1,10).setValue(data.newRole); return { success: true, message: 'Role updated to ' + data.newRole }; }
  }
  return { error: 'User not found' };
}

function sendManualReminder(data) {
  if (!isAdmin(data.requesterUserID)) return { error: 'Unauthorized' };
  var sheet = getSheet(SHEET.BORROWED), rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.borrowID) {
      var user = getUserByID(rows[i][9]);
      if (!user || !user.email) return { error: 'Reader email not found' };
      var due = new Date(rows[i][6]);
      var html = '<div style="font-family:Georgia,serif;max-width:520px;margin:auto;background:#fffbeb;padding:32px;border-radius:14px;border-left:5px solid #c8a44a;"><h2 style="color:#92400e;">📚 अक्षरधारा वाचनालय</h2><p>नमस्कार <strong>' + rows[i][3] + '</strong>,</p><p>आपण घेतलेले पुस्तक <strong>"' + rows[i][2] + '"</strong> परत करण्याची वेळ आली आहे.</p><div style="background:#c8a44a;color:#fff;padding:16px;border-radius:8px;text-align:center;font-size:18px;margin:16px 0;">📅 परत करण्याची तारीख: <strong>' + due.toLocaleDateString('mr-IN') + '</strong></div></div>';
      var r = sendBrevoEmail(user.email, 'अक्षरधारा वाचनालय – पुस्तक परत करण्याची आठवण', html);
      return r.success ? { success: true, message: 'Reminder sent to ' + user.email } : { error: r.error };
    }
  }
  return { error: 'Borrow record not found' };
}

// ============================================================
// DAILY TRIGGER
// ============================================================

function sendDayTenReminders() {
  var rows = getSheet(SHEET.BORROWED).getDataRange().getValues(), now = new Date();
  rows.slice(1).forEach(function(row) {
    if (row[8] !== 'borrowed') return;
    var hrs = (new Date(row[6]) - now) / 3600000;
    if (hrs >= 0 && hrs <= 24) {
      var user = getUserByID(row[9]);
      if (!user || !user.email) return;
      var html = '<div style="font-family:Georgia,serif;max-width:500px;margin:auto;background:#fef2f2;padding:32px;border-radius:14px;border-left:5px solid #8b1e1e;"><h2 style="color:#7f1d1d;">⚠️ अक्षरधारा वाचनालय</h2><p>नमस्कार <strong>' + row[3] + '</strong>,</p><p>आपण घेतलेले पुस्तक <strong>"' + row[2] + '"</strong> आज परत करायचे आहे!</p></div>';
      sendBrevoEmail(user.email, 'अक्षरधारा वाचनालय – आजच पुस्तक परत करा!', html);
    }
  });
}

function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === 'sendDayTenReminders') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendDayTenReminders').timeBased().atHour(9).everyDays(1).create();
  return { success: true, message: 'Daily trigger installed (9 AM)' };
}

// ============================================================
// HELPERS
// ============================================================

function getSheet(name) { var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); if (!s) throw new Error('Sheet "' + name + '" not found. Run setup() first.'); return s; }
function generateID() { return Utilities.getUuid().replace(/-/g,'').substring(0,16); }
function generateOTP() { return Math.floor(100000 + Math.random() * 900000).toString(); }

// SHA-256 password hashing (columns: [5]=AuthProvider, [7]=Password, [9]=Role)
function hashPassword(password) {
  if (!password) return '';
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  return bytes.map(function(b) { return ('0' + (b < 0 ? b+256 : b).toString(16)).slice(-2); }).join('');
}

function rowToUser(row) {
  return { id:row[0], name:row[1], email:row[2], mobile:row[3], age:row[4], authProvider:row[5], isVerified:row[8], role:row[9], createdDate:row[10] };
}

function getUserByID(userID) {
  var rows = getSheet(SHEET.USERS).getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) { if (rows[i][0] === userID) return rowToUser(rows[i]); }
  return null;
}

function isAdmin(userID) { var u = getUserByID(userID); return u && u.role === 'admin'; }

function getConfig() {
  var p = PropertiesService.getScriptProperties();
  return { BREVO_API_KEY: p.getProperty('BREVO_API_KEY')||'', BREVO_SENDER_EMAIL: p.getProperty('BREVO_SENDER_EMAIL')||'', BREVO_SENDER_NAME: p.getProperty('BREVO_SENDER_NAME')||'Akshardhara Vachnalay' };
}

function sendBrevoEmail(to, subject, html) {
  var c = getConfig();
  if (!c.BREVO_API_KEY || !c.BREVO_SENDER_EMAIL) return { error: 'Brevo config missing in Script Properties' };
  try {
    var r = UrlFetchApp.fetch('https://api.brevo.com/v3/smtp/email', { method:'POST', contentType:'application/json', headers:{'api-key':c.BREVO_API_KEY}, payload:JSON.stringify({ sender:{name:c.BREVO_SENDER_NAME,email:c.BREVO_SENDER_EMAIL}, to:[{email:to}], subject:subject, htmlContent:html }), muteHttpExceptions:true });
    var code = r.getResponseCode();
    return (code===200||code===201) ? {success:true} : {error:'Brevo '+code+': '+r.getContentText()};
  } catch(e) { return {error:e.toString()}; }
}
