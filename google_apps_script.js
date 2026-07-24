const SHEET_NAME = 'Reports';
const DRIVE_SHEET_NAME = 'DriveItems';

// ID Thư mục Google Drive của bạn - Tự động lưu file vào đây khi upload từ web
const DRIVE_FOLDER_ID = '17_809JzLyZoYKDP2JDoPhwwTn912HCKQ';

// ★ BẤM CHẠY HÀM NÀY NẾU BỊ BÁO LỖI "Truy cập bị từ chối: DriveApp" ★
function authorizeDrive() {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  Logger.log('Đã kích hoạt quyền Drive thành công cho thư mục: ' + folder.getName());
}

function getOrCreateSheet(sheetName) {
  const name = sheetName || SHEET_NAME;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  
  if (name === 'Reports') {
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['id', 'upload_date', 'title', 'filename', 'data', 'doc_file_url', 'doc_file_id', 'json_file_id']);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    }
  } else if (name === 'DriveItems') {
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['id', 'parentId', 'type', 'name', 'url', 'desc', 'theme', 'date']);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    }
  }
  return sheet;
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'list';

  if (action === 'list') {
    return jsonResponse(listReports());
  } else if (action === 'get') {
    const id = e.parameter.id;
    return jsonResponse(getReport(id));
  } else if (action === 'listDriveItems') {
    return jsonResponse(listDriveItems());
  }
  return jsonResponse({ error: 'Unknown action' });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'save') {
      saveReport(body.report, body.fileData);
      return jsonResponse({ success: true });
    } else if (action === 'delete') {
      deleteReport(body.id);
      return jsonResponse({ success: true });
    } else if (action === 'saveDriveItem') {
      const result = saveDriveItem(body.item);
      return jsonResponse({ success: true, item: result });
    } else if (action === 'deleteDriveItem') {
      deleteDriveItem(body.id);
      return jsonResponse({ success: true });
    }
    return jsonResponse({ error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ─── PHẦN 1: QUẢN LÝ BÁO CÁO TUẦN (LƯU TRỰC TIẾP FILE DOCX VÀ JSON VÀO GOOGLE DRIVE) ───
function listReports() {
  const sheet = getOrCreateSheet('Reports');
  const data = sheet.getDataRange().getValues();
  const reports = [];
  for (let i = 1; i < data.length; i++) {
    reports.push({
      id: data[i][0],
      upload_date: data[i][1],
      title: data[i][2],
      filename: data[i][3],
      doc_file_url: data[i][5] || ''
    });
  }
  // Sắp xếp ngày mới nhất lên đầu
  reports.sort((a, b) => new Date(b.upload_date) - new Date(a.upload_date));
  return reports;
}

function getReport(id) {
  const sheet = getOrCreateSheet('Reports');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      // 1. Ưu tiên đọc từ file JSON lưu trữ trên Google Drive (nếu có json_file_id ở cột 8)
      const jsonFileId = data[i][7];
      if (jsonFileId) {
        try {
          const jsonFile = DriveApp.getFileById(jsonFileId);
          const content = jsonFile.getBlob().getDataAsString();
          return JSON.parse(content);
        } catch (err) {
          Logger.log('Không đọc được file JSON trên Drive: ' + err.message);
        }
      }

      // 2. Tương thích ngược với các báo cáo cũ lưu dạng base64 trong Google Sheet (cột 5)
      if (data[i][4]) {
        try {
          const decoded = Utilities.newBlob(Utilities.base64Decode(data[i][4])).getDataAsString();
          return JSON.parse(decoded);
        } catch (err) {
          return { error: 'Lỗi giải mã báo cáo cũ: ' + err.message };
        }
      }
    }
  }
  return { error: 'Không tìm thấy báo cáo' };
}

function saveReport(report, fileData) {
  const sheet = getOrCreateSheet('Reports');
  let docFileUrl = '';
  let docFileId = '';
  let jsonFileId = '';

  // 1. Lưu file thực tế (.docx gốc) vào Google Drive nếu người dùng gửi fileData
  if (fileData && fileData.indexOf('data:') === 0) {
    try {
      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const parts = fileData.split(',');
      const mime = parts[0].match(/:(.*?);/)[1] || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
      const bytes = Utilities.base64Decode(parts[1]);
      const blob = Utilities.newBlob(bytes, mime, report.filename);
      const docFile = folder.createFile(blob);
      docFileUrl = docFile.getUrl();
      docFileId = docFile.getId();
    } catch (err) {
      Logger.log('Lỗi tạo file .docx trên Drive: ' + err.message);
    }
  }

  // 2. Lưu thông tin dữ liệu đã parse (JSON) thành 1 file .json trên Google Drive để tránh giới hạn ô Sheet
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const jsonBlob = Utilities.newBlob(JSON.stringify(report), 'application/json', (report.filename || report.id) + '.json');
    const jsonFile = folder.createFile(jsonBlob);
    jsonFileId = jsonFile.getId();
  } catch (err) {
    Logger.log('Lỗi tạo file JSON trên Drive: ' + err.message);
  }

  // 3. Lưu thông tin vào Google Sheet (chỉ lưu metadata + id file Drive, không nhồi base64)
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === report.id) {
      sheet.getRange(i + 1, 1, 1, 8).setValues([
        [report.id, report.upload_date, report.title, report.filename, '', docFileUrl || data[i][5], docFileId || data[i][6], jsonFileId || data[i][7]]
      ]);
      return;
    }
  }

  // Thêm dòng mới nếu chưa có
  sheet.appendRow([
    report.id,
    report.upload_date,
    report.title,
    report.filename,
    '', // cột data để trống
    docFileUrl,
    docFileId,
    jsonFileId
  ]);
}

function deleteReport(id) {
  const sheet = getOrCreateSheet('Reports');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      // Xóa file trên Google Drive (nếu có)
      const docFileId = data[i][6];
      const jsonFileId = data[i][7];
      if (docFileId) {
        try { DriveApp.getFileById(docFileId).setTrashed(true); } catch (e) {}
      }
      if (jsonFileId) {
        try { DriveApp.getFileById(jsonFileId).setTrashed(true); } catch (e) {}
      }
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

// ─── PHẦN 2: KHO DỮ LIỆU DÙNG CHUNG ───
function listDriveItems() {
  const sheet = getOrCreateSheet('DriveItems');
  const data = sheet.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < data.length; i++) {
    items.push({
      id: data[i][0],
      parentId: data[i][1],
      type: data[i][2],
      name: data[i][3],
      url: data[i][4],
      desc: data[i][5],
      theme: data[i][6],
      date: data[i][7]
    });
  }
  return items;
}

function saveDriveItem(item) {
  let fileUrl = item.url || '';
  let driveError = '';

  // Nếu URL là data URL (base64 từ máy tính), tạo file thực trên Drive
  if (item.url && item.url.indexOf('data:') === 0) {
    try {
      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const commaIndex = item.url.indexOf(',');
      if (commaIndex === -1) throw new Error('Data URL không hợp lệ');
      
      const header = item.url.substring(0, commaIndex);
      const b64Data = item.url.substring(commaIndex + 1);
      
      const mimeMatch = header.match(/:(.*?);/);
      const mime = mimeMatch ? mimeMatch[1] : 'application/octet-stream';
      
      const bytes = Utilities.base64Decode(b64Data);
      const blob = Utilities.newBlob(bytes, mime, item.name);
      const driveFile = folder.createFile(blob);
      fileUrl = driveFile.getUrl();
      try {
        driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (shareErr) {
        Logger.log('Không thể đặt quyền chia sẻ link công khai: ' + shareErr.message);
      }
      
      Logger.log('Đã tạo file trên Drive: ' + fileUrl + ' (size: ' + bytes.length + ' bytes)');
    } catch (err) {
      driveError = err.message;
      Logger.log('LỖI upload file vào Drive: ' + err.message);
    }
  }

  const sheet = getOrCreateSheet('DriveItems');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === item.id) {
      sheet.getRange(i + 1, 1, 1, 8).setValues([
        [item.id, item.parentId || 'root', item.type || 'file', item.name, fileUrl, item.desc || '', item.theme || '', item.date || new Date().toISOString()]
      ]);
      item.url = fileUrl;
      if (driveError) item.driveError = driveError;
      return item;
    }
  }

  sheet.appendRow([
    item.id,
    item.parentId || 'root',
    item.type || 'file',
    item.name,
    fileUrl,
    item.desc || '',
    item.theme || '',
    item.date || new Date().toISOString()
  ]);

  item.url = fileUrl;
  if (driveError) item.driveError = driveError;
  return item;
}

function deleteDriveItem(id) {
  const sheet = getOrCreateSheet('DriveItems');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
