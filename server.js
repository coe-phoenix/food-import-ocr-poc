const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

const upload = multer({ storage: multer.memoryStorage() });

app.get('/', (_req, res) => res.sendFile(path.join(__dirname, 'index.html')));

app.post('/upload', upload.single('image'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No image uploaded' });

    const extractedData = {
      supplierName: 'Pacific Seafoods Pte Ltd',
      deliveryDate: new Date().toISOString().split('T')[0],
      poNumber: 'PO-2024-' + Math.floor(1000 + Math.random() * 9000),
      items: [
        { item: 'Imported Norwegian Salmon (Fillet)', quantity: 50, unit: 'kg', unitPrice: 28.50, totalPrice: 1425.00 },
        { item: 'Japanese Wagyu Beef A5 (Striploin)',  quantity: 20, unit: 'kg', unitPrice: 145.00, totalPrice: 2900.00 },
        { item: 'Italian Truffle Oil (Extra Virgin)',   quantity: 30, unit: 'btl', unitPrice: 42.00, totalPrice: 1260.00 },
        { item: 'Spanish Saffron (Grade 1)',            quantity: 5,  unit: 'g',   unitPrice: 18.00, totalPrice: 90.00 },
        { item: 'French Dijon Mustard (Maille)',        quantity: 40, unit: 'jar', unitPrice: 12.50, totalPrice: 500.00 }
      ]
    };

    const grandTotal = extractedData.items.reduce((s, i) => s + i.totalPrice, 0);

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Delivery Order');

    ws.mergeCells('A1:F1');
    ws.getCell('A1').value = 'DELIVERY ORDER — DATA EXTRACTION';
    ws.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FF1A73E8' } };
    ws.getCell('A1').alignment = { horizontal: 'center' };

    ws.getCell('A3').value = 'Supplier:';
    ws.getCell('B3').value = extractedData.supplierName;
    ws.getCell('B3').font = { bold: true };
    ws.getCell('A4').value = 'Delivery Date:';
    ws.getCell('B4').value = extractedData.deliveryDate;
    ws.getCell('A5').value = 'PO Number:';
    ws.getCell('B5').value = extractedData.poNumber;
    ws.getCell('B5').font = { bold: true };
    ws.getCell('A6').value = 'Source File:';
    ws.getCell('B6').value = req.file.originalname;

    ws.addRow([]);
    const tableHeader = ws.addRow(['#', 'Item Description', 'Qty', 'Unit', 'Unit Price (USD)', 'Total Price (USD)']);
    tableHeader.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    tableHeader.eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A73E8' } };
      cell.alignment = { horizontal: 'center' };
      cell.border = { bottom: { style: 'thin' } };
    });

    extractedData.items.forEach((item, idx) => {
      const row = ws.addRow([idx + 1, item.item, item.quantity, item.unit, item.unitPrice, item.totalPrice]);
      row.getCell(5).numFmt = '#,##0.00';
      row.getCell(6).numFmt = '#,##0.00';
      if (idx % 2 === 0) {
        row.eachCell(cell => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F6FC' } };
        });
      }
    });

    const totalRow = ws.addRow(['', '', '', '', 'GRAND TOTAL', grandTotal]);
    totalRow.font = { bold: true, size: 12 };
    totalRow.getCell(6).numFmt = '#,##0.00';
    totalRow.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE082' } };

    ws.columns = [
      { width: 5 }, { width: 38 }, { width: 8 }, { width: 8 }, { width: 18 }, { width: 20 }
    ];

    const buffer = await wb.xlsx.writeBuffer();
    const filename = `delivery-order-${extractedData.poNumber}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=\"${filename}\"`);
    res.send(buffer);

  } catch (err) {
    console.error('Upload error:', err);
    res.status(500).json({ error: 'Processing failed' });
  }
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
