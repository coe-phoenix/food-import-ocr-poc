require('dotenv').config();
const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const AnthropicModule = require('@anthropic-ai/sdk');
const path = require('path');

const Anthropic = AnthropicModule.default || AnthropicModule;

const app = express();
const PORT = process.env.PORT || 3000;

const anthropic = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY,
});

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 },
});

// ---------- Serve frontend ----------
app.get('/', (_req, res) => res.sendFile(path.join(__dirname, 'index.html')));

// ---------- Claude Vision extraction ----------
async function extractDeliveryOrder(fileBuffer, mimeType) {
  const base64Data = fileBuffer.toString('base64');

  // Build the content block based on file type
  let fileContent;
  if (mimeType === 'application/pdf') {
    fileContent = {
      type: 'document',
      source: {
        type: 'base64',
        media_type: 'application/pdf',
        data: base64Data,
      },
    };
  } else {
    fileContent = {
      type: 'image',
      source: {
        type: 'base64',
        media_type: mimeType,
        data: base64Data,
      },
    };
  }

  const response = await anthropic.messages.create({
    model: 'claude-sonnet-4-20250514',
    max_tokens: 4096,
    messages: [
      {
        role: 'user',
        content: [
          fileContent,
          {
            type: 'text',
            text: `You are a data extraction assistant for a food import company.

Analyze this delivery order / invoice document and extract ALL data you can find.

Return ONLY valid JSON in this exact format (no markdown, no backticks, no explanation):
{
  "supplierName": "<supplier/vendor name or 'Unknown'>",
  "deliveryDate": "<date in YYYY-MM-DD format or today's date if not found>",
  "poNumber": "<PO/invoice/DO number or 'N/A'>",
  "currency": "<currency code like USD, MYR, SGD, etc. or 'USD' if not found>",
  "items": [
    {
      "item": "<item description>",
      "quantity": <number>,
      "unit": "<unit like kg, pcs, btl, box, etc.>",
      "unitPrice": <number>,
      "totalPrice": <number>
    }
  ]
}

Rules:
- Extract every line item you can see in the document.
- If a field is not visible, make your best reasonable inference from context.
- If unit price is missing but total and quantity exist, calculate it.
- If total price is missing but unit price and quantity exist, calculate it.
- Numbers must be plain numbers (no currency symbols).
- Return raw JSON only.`,
          },
        ],
      },
    ],
  });

  const raw = response.content[0].text.trim();
  const cleaned = raw.replace(/^```(?:json)?\n?/i, '').replace(/\n?```$/i, '').trim();

  return JSON.parse(cleaned);
}

// ---------- POST /upload ----------
app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const mimeType = req.file.mimetype;
    const allowedTypes = ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'application/pdf'];

    if (!allowedTypes.includes(mimeType)) {
      return res.status(400).json({
        error: `Unsupported file type: ${mimeType}. Use JPG, PNG, GIF, WebP, or PDF.`,
      });
    }

    console.log(`Processing: ${req.file.originalname} (${(req.file.size / 1024).toFixed(1)} KB, ${mimeType})`);
    const extractedData = await extractDeliveryOrder(req.file.buffer, mimeType);
    console.log(`Extracted ${extractedData.items.length} line items from ${req.file.originalname}`);

    const grandTotal = extractedData.items.reduce((s, i) => s + (i.totalPrice || 0), 0);
    const currency = extractedData.currency || 'USD';

    // --- Build Excel workbook ---
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Delivery Order');

    ws.mergeCells('A1:F1');
    ws.getCell('A1').value = 'DELIVERY ORDER — AI EXTRACTED DATA';
    ws.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FF1A73E8' } };
    ws.getCell('A1').alignment = { horizontal: 'center' };

    ws.getCell('A3').value = 'Supplier:';
    ws.getCell('B3').value = extractedData.supplierName || 'Unknown';
    ws.getCell('B3').font = { bold: true };
    ws.getCell('A4').value = 'Delivery Date:';
    ws.getCell('B4').value = extractedData.deliveryDate || new Date().toISOString().split('T')[0];
    ws.getCell('A5').value = 'PO Number:';
    ws.getCell('B5').value = extractedData.poNumber || 'N/A';
    ws.getCell('B5').font = { bold: true };
    ws.getCell('A6').value = 'Source File:';
    ws.getCell('B6').value = req.file.originalname;
    ws.getCell('A7').value = 'Extracted By:';
    ws.getCell('B7').value = 'Claude Vision AI';
    ws.getCell('B7').font = { italic: true, color: { argb: 'FF6B7280' } };

    ws.addRow([]);
    const tableHeader = ws.addRow(['#', 'Item Description', 'Qty', 'Unit', `Unit Price (${currency})`, `Total Price (${currency})`]);
    tableHeader.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    tableHeader.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A73E8' } };
      cell.alignment = { horizontal: 'center' };
      cell.border = { bottom: { style: 'thin' } };
    });

    extractedData.items.forEach((item, idx) => {
      const row = ws.addRow([
        idx + 1,
        item.item || '',
        item.quantity || 0,
        item.unit || '',
        item.unitPrice || 0,
        item.totalPrice || 0,
      ]);
      row.getCell(5).numFmt = '#,##0.00';
      row.getCell(6).numFmt = '#,##0.00';
      if (idx % 2 === 0) {
        row.eachCell((cell) => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F6FC' } };
        });
      }
    });

    const totalRow = ws.addRow(['', '', '', '', 'GRAND TOTAL', grandTotal]);
    totalRow.font = { bold: true, size: 12 };
    totalRow.getCell(6).numFmt = '#,##0.00';
    totalRow.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE082' } };

    ws.columns = [
      { width: 5 },
      { width: 40 },
      { width: 8 },
      { width: 8 },
      { width: 18 },
      { width: 18 },
    ];

    const buffer = await wb.xlsx.writeBuffer();
    const safePoNumber = (extractedData.poNumber || 'extracted').replace(/[^a-zA-Z0-9-_]/g, '-');
    const filename = `delivery-order-${safePoNumber}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch (err) {
    console.error('Upload error:', err);

    if (err.status === 401) {
      return res.status(500).json({ error: 'Invalid Anthropic API key. Check your .env file.' });
    }
    if (err instanceof SyntaxError) {
      return res.status(500).json({ error: 'AI returned unparseable data. Try a clearer image.' });
    }

    res.status(500).json({ error: 'Processing failed: ' + (err.message || 'Unknown error') });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  if (!process.env.ANTHROPIC_API_KEY) {
    console.warn('\n⚠️  WARNING: ANTHROPIC_API_KEY is not set!');
    console.warn('   Create a .env file with: ANTHROPIC_API_KEY=sk-ant-...');
    console.warn('   Without it, image extraction will fail.\n');
  }
});
