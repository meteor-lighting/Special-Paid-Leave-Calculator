import * as XLSX from 'xlsx';

export const exportToExcel = (data, filename = 'Special_Paid_Leave_Report.xlsx') => {
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Report');
  XLSX.writeFile(workbook, filename);
};

export const exportPersonDetails = (personName, year, records) => {
  const formattedData = records.map(r => ({
    '開始日期': r.startDate,
    '結束日期': r.endDate,
    '出差地點': r.location || '',
    '出差天數': r.tripDays,
    '換算週數': r.weeks,
    '本次產生天數': r.generatedLeave,
    '前次小數餘額': r.prevRemainder,
    '本次核給天數': r.awarded,
    '剩餘小數餘額': r.remainder,
    '備註': r.note || ''
  }));

  const worksheet = XLSX.utils.json_to_sheet(formattedData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Details');
  XLSX.writeFile(workbook, `${personName}_${year}_Leave_Details.xlsx`);
};
