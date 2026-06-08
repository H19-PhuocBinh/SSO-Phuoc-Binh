const SHEET_NAME = 'HeThong';

const HEADERS = [
  'Nhóm',
  'Tên hệ thống',
  'Link',
  'Icon',
  'Màu 1',
  'Màu 2',
  'Thứ tự',
  'Hiển thị',
];

const SAMPLE_ROWS = [
  ['Dùng chung', 'Hệ thống Quản lý văn bản', 'https://dnis.dongnai.gov.vn/', 'document', '#1976d2', '#42a5f5', 1, 'Có'],
  ['Dùng chung', 'Hệ thống Họp không giấy', 'https://hop.dongnai.gov.vn/', 'meeting', '#16a34a', '#4ade80', 2, 'Có'],
  ['Quản trị', 'Hệ thống Quản trị dữ liệu TTHC', 'https://quantri.dichvucong.gov.vn/', 'database', '#7c3aed', '#a78bfa', 3, 'Có'],
  ['Quản trị', 'Hệ thống CSDL TTHC', 'https://csdltthc.dichvucong.gov.vn/sso', 'database', '#f97316', '#fb923c', 4, 'Có'],
  ['Quản trị', 'Hệ thống Quản trị TTTT', 'https://quantrithanhtoan.ndc.gov.vn/', 'payment', '#2563eb', '#60a5fa', 5, 'Có'],
  ['Quản trị', 'Hệ thống Quản trị DVC', 'https://quantricong.dichvucong.gov.vn/sso', 'service', '#0891b2', '#22d3ee', 6, 'Có'],
  ['Quản trị', 'Hệ thống Phản ánh kiến nghị', 'https://phananhkiennghi.dichvucong.gov.vn/', 'feedback', '#db2777', '#fb7185', 7, 'Có'],
  ['Một cửa', 'Cổng Dịch vụ công', 'https://dichvucong.gov.vn/', 'portal', '#b40000', '#facc15', 8, 'Có'],
  ['Một cửa', 'Hệ thống iGate thành phố Đồng Nai', 'https://motcua.dongnai.gov.vn/', 'cityGate', '#0f766e', '#14b8a6', 9, 'Có'],
  ['Một cửa', 'Bộ Khoa học và Công nghệ', 'https://motcua.mst.gov.vn/', 'science', '#7c3aed', '#a855f7', 10, 'Có'],
  ['Một cửa', 'Bộ Nội Vụ', 'https://motcua.moha.gov.vn/', 'people', '#2563eb', '#60a5fa', 11, 'Có'],
  ['Một cửa', 'Bộ Xây dựng', 'https://motcuabxd.moc.gov.vn/', 'construction', '#f97316', '#fbbf24', 12, 'Có'],
  ['Một cửa', 'Bộ Công thương', 'https://motcua-tthc.moit.gov.vn/', 'industry', '#0f766e', '#2dd4bf', 13, 'Có'],
  ['Một cửa', 'Bộ Nông nghiệp và Môi Trường', 'https://motcuannmt.mae.gov.vn/', 'environment', '#16a34a', '#86efac', 14, 'Có'],
  ['Một cửa', 'Bộ Y tế', 'https://motcua.moh.gov.vn/', 'health', '#dc2626', '#fb7185', 15, 'Có'],
  ['Một cửa', 'Bộ Giáo dục và Đào tạo', 'https://motcua.moet.gov.vn/', 'education', '#1d4ed8', '#38bdf8', 16, 'Có'],
  ['Một cửa', 'Bộ Dân tộc và Tôn giáo', 'https://gqtthc.bdttg.gov.vn/', 'people', '#b45309', '#f59e0b', 17, 'Có'],
  ['Một cửa', 'Bộ Tư pháp', 'https://motcua.moj.gov.vn/', 'justice', '#4338ca', '#818cf8', 18, 'Có'],
  ['Một cửa', 'Bộ Ngoại giao', 'https://bpmc.mofa.gov.vn/', 'foreign', '#0284c7', '#38bdf8', 19, 'Có'],
  ['Một cửa', 'Bộ Tài chính', 'https://dichvucong.mof.gov.vn/', 'finance', '#15803d', '#4ade80', 20, 'Có'],
  ['Một cửa', 'Bộ Văn hóa, Thể thao và Du lịch', 'https://dichvucong.bvhttdl.gov.vn/tiepnhan', 'culture', '#be123c', '#fb7185', 21, 'Có'],
];

function setupSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);

  sheet.clear();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.getRange(2, 1, SAMPLE_ROWS.length, HEADERS.length).setValues(SAMPLE_ROWS);

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, HEADERS.length);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setBackground('#b40000')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
}

function doGet() {
  const data = getSystems();

  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      updatedAt: new Date().toISOString(),
      groups: countByGroup(data),
      data,
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSystems() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    return [];
  }

  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1);

  return rows
    .map((row) => ({
      group: String(row[0] || '').trim(),
      name: String(row[1] || '').trim(),
      url: String(row[2] || '').trim(),
      icon: String(row[3] || 'portal').trim(),
      colors: [
        String(row[4] || '#b40000').trim(),
        String(row[5] || '#facc15').trim(),
      ],
      order: Number(row[6] || 9999),
      visible: String(row[7] || 'Có').trim().toLowerCase() !== 'không',
    }))
    .filter((item) => item.visible && item.name && /^https?:\/\//i.test(item.url))
    .sort((a, b) => a.order - b.order);
}

function countByGroup(data) {
  return data.reduce((result, item) => {
    result[item.group] = (result[item.group] || 0) + 1;
    return result;
  }, {});
}
