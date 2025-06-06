import Company from '../models/Company.js';
import ExcelJS from 'exceljs';

export const getAllCompanies = async (req, res) => {
  const companies = await Company.find();
  res.json(companies);
};

export const getCompanyById = async (req, res) => {
  const company = await Company.findById(req.params.id);
  res.json(company);
};

export const createCompany = async (req, res) => {
  try {
    const containerType = req.body.containerType;
    const containerQuantity = Number(req.body.containerQuantity) || 0;
    const portAuthorityFee = Number(req.body.portAuthorityFee) || 0;
    const seaportFee = Number(req.body.seaportFee) || 0;
    const emptyPortFee = Number(req.body.emptyPortFee) || 0;
    const unloadingFee = Number(req.body.unloadingFee) || 0;
    const warehouseFee = Number(req.body.warehouseFee) || 0;
    const directDeliveryFee = Number(req.body.directDeliveryFee) || 0;
    const hiepPhuocPortFee = Number(req.body.hiepPhuocPortFee) || 0;
    const serviceFee = Number(req.body.serviceFee) || 0;
    const deliveryServiceFee = Number(req.body.deliveryServiceFee) || 0;
    const customsFee = Number(req.body.customsFee) || 0;
    const liftingFee = Number(req.body.liftingFee) || 0;
    const botFee = 160000 * containerQuantity;

    let transportFee = 0;
    if (containerType === '20ft') {
      transportFee = 2700000 * containerQuantity;
    } else if (containerType === '40ft') {
      transportFee = 3400000 * containerQuantity;
    } else {
      transportFee = 700000 * containerQuantity;
    }

    const serviceTotal10 = deliveryServiceFee + customsFee + liftingFee + transportFee + hiepPhuocPortFee + botFee;
    const serviceTotal8 = (serviceTotal10 / 1.1) * 1.08;

    const totalAmount = portAuthorityFee + seaportFee + emptyPortFee + warehouseFee + directDeliveryFee + serviceFee + serviceTotal8 + unloadingFee;

    const newCompany = new Company({
      ...req.body,
      containerType,
      containerQuantity,
      portAuthorityFee,
      seaportFee,
      emptyPortFee,
      transportFee,
      warehouseFee,
      directDeliveryFee,
      hiepPhuocPortFee,
      serviceFee,
      deliveryServiceFee,
      customsFee,
      liftingFee,
      botFee,
      serviceTotal10,
      serviceTotal8,
      unloadingFee,
      totalAmount
    });

    const saved = await newCompany.save();
    res.json(saved);
  } catch (error) {
    res.status(500).json({ message: 'Lỗi khi tạo công ty', error });
  }
};

export const updateCompany = async (req, res) => {
  try {
    const containerType = req.body.containerType;
    const containerQuantity = Number(req.body.containerQuantity) || 0;
    const portAuthorityFee = Number(req.body.portAuthorityFee) || 0;
    const seaportFee = Number(req.body.seaportFee) || 0;
    const emptyPortFee = Number(req.body.emptyPortFee) || 0;
    const unloadingFee = Number(req.body.unloadingFee) || 0;
    const warehouseFee = Number(req.body.warehouseFee) || 0;
    const directDeliveryFee = Number(req.body.directDeliveryFee) || 0;
    const hiepPhuocPortFee = Number(req.body.hiepPhuocPortFee) || 0;
    const serviceFee = Number(req.body.serviceFee) || 0;
    const deliveryServiceFee = Number(req.body.deliveryServiceFee) || 0;
    const customsFee = Number(req.body.customsFee) || 0;
    const liftingFee = Number(req.body.liftingFee) || 0;
    const botFee = 160000 * containerQuantity;

    let transportFee = 0;
    if (containerType === '20ft') {
      transportFee = 2700000 * containerQuantity;
    } else if (containerType === '40ft') {
      transportFee = 3400000 * containerQuantity;
    } else {
      transportFee = 700000 * containerQuantity;
    }

    const serviceTotal10 = deliveryServiceFee + customsFee + liftingFee + transportFee + hiepPhuocPortFee + botFee;
    const serviceTotal8 = (serviceTotal10 / 1.1) * 1.08;

    const totalAmount = portAuthorityFee + seaportFee + emptyPortFee + warehouseFee + directDeliveryFee + serviceFee + serviceTotal8 + unloadingFee;

    const updated = await Company.findByIdAndUpdate(
      req.params.id,
      {
        ...req.body,
        containerType,
        containerQuantity,
        portAuthorityFee,
        seaportFee,
        emptyPortFee,
        transportFee,
        warehouseFee,
        directDeliveryFee,
        hiepPhuocPortFee,
        serviceFee,
        deliveryServiceFee,
        customsFee,
        liftingFee,
        botFee,
        serviceTotal10,
        serviceTotal8,
        unloadingFee,
        totalAmount
      },
      { new: true }
    );

    res.json(updated);
  } catch (error) {
    res.status(500).json({ message: 'Lỗi khi cập nhật công ty', error });
  }
};

export const deleteCompany = async (req, res) => {
  await Company.findByIdAndDelete(req.params.id);
  res.json({ message: 'Deleted' });
};

export const exportCompanyExcel = async (req, res) => {
  const { companyName } = req.params;
  try {
    const entries = await Company.find({ companyName });
    if (!entries.length) {
      return res.status(404).json({ message: 'Không tìm thấy dữ liệu công ty' });
    }

    const workbook = new ExcelJS.Workbook();
    const safeSheetName = companyName.replace(/[\\/*?:\[\]]/g, '_');
    const worksheet = workbook.addWorksheet("Danh sách - " + safeSheetName);

    // Tiêu đề chính
    worksheet.mergeCells('A1:U1');
    worksheet.getCell('A1').value = `CÔNG NỢ - ${companyName}`;
    worksheet.getCell('A1').font = { size: 14, bold: true };
    worksheet.getCell('A1').alignment = { horizontal: 'center' };

    // Tiêu đề dòng 2
    worksheet.getCell('A2').value = 'Ngày khai HQ';
    worksheet.getCell('A2').alignment = { horizontal: 'center' };
    worksheet.getCell('A2').font = { bold: true, size: 7 };
    worksheet.getCell('B2').value = 'Số tờ khai';
    worksheet.getCell('B2').alignment = { horizontal: 'center' };
    worksheet.getCell('B2').font = { bold: true, size: 7 };
    worksheet.getCell('C2').value = 'Mã hợp đồng';
    worksheet.getCell('C2').alignment = { horizontal: 'center' };
    worksheet.getCell('C2').font = { bold: true, size: 7 };
    worksheet.getCell('D2').value = 'Phương tiện vận chuyển';
    worksheet.getCell('DD2').alignment = { horizontal: 'center' };
    worksheet.getCell('D2').font = { bold: true, size: 7 };

    worksheet.mergeCells('E2:K2');
    worksheet.getCell('E2').value = 'Phí thu hộ';
    worksheet.getCell('E2').alignment = { horizontal: 'center' };
    worksheet.getCell('E2').font = { bold: true, size: 7 };

    worksheet.mergeCells('L2:S2');
    worksheet.getCell('L2').value = 'Phí dịch vụ';
    worksheet.getCell('L2').alignment = { horizontal: 'center' };
    worksheet.getCell('L2').font = { bold: true, size: 7 };

    worksheet.getCell('T2').value = 'Thành tiền';
    worksheet.getCell('T2').alignment = { horizontal: 'center' };
    worksheet.getCell('T2').font = { bold: true, size: 7 };
    worksheet.getCell('U2').value = 'Ghi chú';
    worksheet.getCell('U2').alignment = { horizontal: 'center' };
    worksheet.getCell('U2').font = { bold: true, size: 7 };

    // Áp dụng font size 9 cho các cột A đến D từ dòng 4 trở đi
    ['A','B','C','D'].forEach(col => {
    worksheet.getColumn(col).eachCell({ includeEmpty: true }, (cell, rowNumber) => {
        if (rowNumber > 3) cell.font = { size: 9 };
      });
   });


    // Dòng tiêu đề con - dòng 3
    const headersRow3 = [
      'Phí cảng vụ', 'Phí cảng biển', 'Phí cảng rỗng',
      'Phí vận chuyển về kho', 'Phí vận chuyển giao thẳng', 'Phí dỡ hàng', 'Tổng phí thu hộ',
      'Phí vận chuyển', 'Phí cảng Hiệp Phước', 'Dịch vụ giao nhận',
      'Hạ trái tuyến', 'Phí BOT', 'Lệ phí HQ', 'Tổng phí DV 10%', 'Tổng phí DV 8%'
    ];

    headersRow3.forEach((title, index) => {
      const col = String.fromCharCode('E'.charCodeAt(0) + index);
      worksheet.getCell(`${col}3`).value = title;
      worksheet.getCell(`${col}3`).alignment = { horizontal: 'center' };
      worksheet.getCell(`${col}3`).font = { bold: true, size: 7 };
    });

    entries.forEach(entry => {
      const formattedDate = new Date(entry.date).toLocaleDateString('vi-VN', {
        day: '2-digit', month: '2-digit', year: '2-digit'
      }).replace(/\//g, '.');

      const row = worksheet.addRow([
        formattedDate,
        `${entry.declarationNumber || ''}\n${entry.containerQuantity || ''}x${entry.containerType || ''}\n${entry.packageCount || ''} kiện/${entry.weightKg || ''}kg`,
        entry.contractCode || '',
        entry.transport || '',

        entry.portAuthorityFee || 0,
        entry.seaportFee || 0,
        entry.emptyPortFee || 0,
        entry.warehouseFee || 0,
        entry.directDeliveryFee || 0,
        entry.unloadingFee || 0,
        (entry.portAuthorityFee || 0) + (entry.seaportFee || 0) + (entry.emptyPortFee || 0) + (entry.warehouseFee || 0) + (entry.directDeliveryFee || 0) + (entry.unloadingFee || 0),

        entry.transportFee || 0,
        entry.hiepPhuocPortFee || 0,
        entry.deliveryServiceFee || 0,
        entry.liftingFee || 0,
        entry.botFee || 0,
        entry.customsFee || 0,
        entry.serviceTotal10 || 0,
        entry.serviceTotal8 || 0,

        entry.totalAmount || 0,
        entry.note || ''
      ]);

      row.getCell(2).alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
      row.getCell(2).font = { size: 7 };
      row.getCell(3).alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
      row.getCell(3).font = { size: 7 };


      // Định dạng các cột số với dấu phẩy hàng nghìn
      [5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20].forEach(index => {
        const cell = row.getCell(index);
        cell.font = { size: 9 };
        if (!isNaN(cell.value)) {
          cell.numFmt = '#,##0';
          cell.alignment = { horizontal: 'right' };
        }
      });
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${safeSheetName}.xlsx`);
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    res.status(500).json({ message: 'Lỗi khi xuất file Excel', error });
  }
};

