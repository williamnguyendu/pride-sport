const SHEET_ID = '1Mu25JGgxcaKwMZ2iklVcb9wvEVibR4QW-WUIxkQl-ac';

function doGet(e) {
  if (e.parameter.action === 'getEvent') {
    return getEventData();
  }
  return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid action' })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const hoVaTen = e.parameter.hoVaTen;
    let soDienThoai = e.parameter.soDienThoai;

    // Chuẩn hóa số điện thoại: Thêm số 0 ở đầu nếu cần
    soDienThoai = soDienThoai.replace(/\D/g, ''); 
    if (soDienThoai.startsWith('84')) {
      soDienThoai = '0' + soDienThoai.slice(2);
    }
    if (!soDienThoai.startsWith('0')) {
      soDienThoai = '0' + soDienThoai;
    }

    // Thêm dấu nháy đơn trước số điện thoại để Google Sheets giữ nguyên số 0
    const soDienThoaiText = "'" + soDienThoai;

    // Format thời gian: Ngày/Tháng/Năm Giờ:Phút:Giây
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DangKySport');
    sheet.appendRow([timestamp, hoVaTen, soDienThoaiText]);

    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getEventData() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('LichSport');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) throw new Error("Chưa có dữ liệu sự kiện");
    
    const headers = data[0];
    const rows = data.slice(1).map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });

    // Lọc sự kiện "Mở đăng ký"
    let activeEvents = rows.filter(row => row.TrangThai === 'Mở đăng ký');
    
    if (activeEvents.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Hiện không có sự kiện nào mở đăng ký.' })).setMimeType(ContentService.MimeType.JSON);
    }

    // Sắp xếp lấy ngày tổ chức lớn nhất nếu có nhiều dòng Mở đăng ký
    activeEvents.sort((a, b) => new Date(b.NgayToChuc).getTime() - new Date(a.NgayToChuc).getTime());
    const event = activeEvents[0];

    // Trả về JSON
    const responseData = {
      TenSport: event.TenSport,
      NgayToChuc: Utilities.formatDate(new Date(event.NgayToChuc), "GMT+7", "dd/MM/yyyy"),
      MonToChuc: event.MonToChuc,
      DiaDiem: event.DiaDiem,
      DangKyTu: new Date(event.DangKyTu).getTime(), // Trả về timestamp để Front-end dễ so sánh
      DangKyTuStr: Utilities.formatDate(new Date(event.DangKyTu), "GMT+7", "dd/MM/yyyy HH:mm"),
      DangKyDen: new Date(event.DangKyDen).getTime(),
      DangKyDenStr: Utilities.formatDate(new Date(event.DangKyDen), "GMT+7", "dd/MM/yyyy HH:mm"),
      LuuY: event.LuuY
    };

    return ContentService.createTextOutput(JSON.stringify({ success: true, data: responseData }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}