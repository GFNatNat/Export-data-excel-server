import ExcelJS from 'exceljs';

const labelMap = {
  date: 'Ngày khai HQ',
  companyName: 'Tên công ty',
  declarationNumber: 'Số tờ khai',
  packageCount: 'Số kiện',
  weightKg: 'Số kg',
  transport: 'Phương tiện vận chuyển',
  containerType: 'LoạiLoại container',
  containerQuantity: 'Số lượng container',
  portAuthorityFee: 'Phí cảng vụ',
  seaportFee: 'Phí cảng biển',
  emptyPortFee: 'Phí cảng rỗng',
  transportFee: 'Phí vận chuyển',
  warehouseFee: 'Phí vận chuyển về kho',
  directDeliveryFee: 'Phí vận chuyển giao thẳng',
  hiepPhuocPortFee: 'Phí cảng Hiệp Phước',
  serviceFee: 'Phí dịch vụ',
  totalAmount: 'Thành tiền',
  note: 'Ghi chú',
  contractCode: 'Mã hợp đồng'
};

const keys = Object.keys(labelMap);

export const generateExcel = async (company) => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Dashboard');
  sheet.addRow(keys.map(k => labelMap[k]));
  sheet.addRow(keys.map(k => company[k] ?? ''));
  return workbook;
};

export const generateGroupedExcelByCompany = async (companies) => {
  const workbook = new ExcelJS.Workbook();

  const grouped = companies.reduce((acc, item) => {
    const name = item.companyName || 'Unknown';
    if (!acc[name]) acc[name] = [];
    acc[name].push(item);
    return acc;
  }, {});

  for (const companyName in grouped) {
    const sheet = workbook.addWorksheet(companyName);
    sheet.addRow(keys.map(k => labelMap[k]));
    grouped[companyName].forEach((item) => {
      sheet.addRow(keys.map(k => item[k] ?? ''));
    });
  }

  return workbook;
};

export const generateExcelForCompanyName = async (companyName, entries) => {
  const workbook = new ExcelJS.Workbook();

  const safeSheetName = companyName.replace(/[\\/*?:\[\]]/g, '_');
  const worksheet = workbook.addWorksheet("Danh sách - " + safeSheetName);

  worksheet.columns = [
    { header: 'Ngày khai HQ', key: 'date', width: 15 },
      { header: 'Tên công ty', key: 'companyName', width: 20 },
      { header: 'Số tờ khai', key: 'declarationNumber', width: 15 },
      { header: 'Số kiện', key: 'packageCount', width: 10 },
      { header: 'Số kg', key: 'weightKg', width: 10 },
      { header: 'Phương tiện vận chuyển', key: 'transport', width: 20 },
      { header: 'Loại container', key: 'containerType', width: 15 },
      { header: 'Số lượng container', key: 'containerQuantity', width: 20 },
      { header: 'Phí cảng vụ', key: 'portAuthorityFee', width: 15 },
      { header: 'Phí cảng biển', key: 'seaportFee', width: 15 },
      { header: 'Phí cảng rỗng', key: 'emptyPortFee', width: 18 },
      { header: 'Phí vận chuyển', key: 'transportFee', width: 18 },
      { header: 'Phí vận chuyển về kho', key: 'warehouseFee', width: 22 },
      { header: 'Phí vận chuyển giao thẳng', key: 'directDeliveryFee', width: 26 },
      { header: 'Phí cảng Hiệp Phước', key: 'hiepPhuocPortFee', width: 22 },
      { header: 'Dịch vụ giao nhận', key: 'deliveryServiceFee', width: 20 },
      { header: 'Lệ phí HQ', key: 'customsFee', width: 15 },
      { header: 'Hạ trái tuyến', key: 'liftingFee', width: 18 },
      { header: 'BOT', key: 'botFee', width: 15 },
      { header: 'Tổng phí DV 10%', key: 'serviceTotal10', width: 20 },
      { header: 'Tổng phí DV 8%', key: 'serviceTotal8', width: 20 },
      { header: 'Thành tiền', key: 'totalAmount', width: 15 },
      { header: 'Ghi chú', key: 'note', width: 20 },
      { header: 'Mã hợp đồng', key: 'contractCode', width: 20 }
  ];

  entries.forEach(entry => worksheet.addRow(entry));
  return workbook;
};
