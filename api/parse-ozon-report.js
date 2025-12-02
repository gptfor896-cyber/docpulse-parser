export const config = {
  runtime: "nodejs", // важная строка – говорим Vercel использовать Node.js, а не Edge
};

import ExcelJS from "exceljs";

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({
        ok: false,
        error: "Use POST",
      });
    }

    const { fileUrl } = req.body;

    if (!fileUrl) {
      return res.status(400).json({
        ok: false,
        error: "Missing fileUrl",
      });
    }

    // 1. Скачиваем файл (используем встроенный fetch, БЕЗ node-fetch)
    const response = await fetch(fileUrl);
    if (!response.ok) {
      return res.status(400).json({
        ok: false,
        error: `Failed to download file: ${response.status}`,
      });
    }

    const arrayBuffer = await response.arrayBuffer();

    // 2. Читаем Excel
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(Buffer.from(arrayBuffer));

    const sheet = workbook.getWorksheet(1); // Первый лист

    let operations = [];

    // 2.1. Находим строку заголовков (ищем ячейку с текстом "Артикул продавца")
    let headerRowIndex = null;

    sheet.eachRow((row, rowNumber) => {
      const text = row.values
        .map((v) => (typeof v === "string" ? v : ""))
        .join(" ");
      if (text.includes("Артикул продавца")) {
        headerRowIndex = rowNumber;
      }
    });

    if (!headerRowIndex) {
      return res.status(400).json({
        ok: false,
        error: "Header row not found (no 'Артикул продавца')",
      });
    }

    const headerRow = sheet.getRow(headerRowIndex);
    const headers = headerRow.values.map((h) =>
      typeof h === "string" ? h.trim() : ""
    );

    const col = (name) => headers.indexOf(name);

    const colSku = col("Артикул продавца");
    const colQtySale = col("Количество");
    const colAmountSale = col("Итого к начислению, руб.");
    const colQtyReturn = col("Количество возвратов");
    const colAmountReturn = col("Итого возвращено, руб.");

    // 3. Идём по строкам ниже заголовков
    for (let i = headerRowIndex + 1; i <= sheet.rowCount; i++) {
      const row = sheet.getRow(i);
      if (!row || !row.values) continue;

      const sku = colSku > 0 ? row.getCell(colSku).value : null;
      if (!sku) continue; // пустая строка – пропускаем

      const qtySale = colQtySale > 0 ? row.getCell(colQtySale).value || 0 : 0;
      const amountSale =
        colAmountSale > 0 ? row.getCell(colAmountSale).value || 0 : 0;

      const qtyReturn =
        colQtyReturn > 0 ? row.getCell(colQtyReturn).value || 0 : 0;
      const amountReturn =
        colAmountReturn > 0 ? row.getCell(colAmountReturn).value || 0 : 0;

      // Продажа
      if (amountSale && Number(amountSale) !== 0) {
        operations.push({
          operation_type: "sale",
          sku,
          quantity: Number(qtySale) || 0,
          amount: Number(amountSale) || 0,
        });
      }

      // Возврат
      if (amountReturn && Number(amountReturn) !== 0) {
        operations.push({
          operation_type: "return",
          sku,
          quantity: Number(qtyReturn) || 0,
          amount: -Math.abs(Number(amountReturn) || 0),
        });
      }
    }

    return res.status(200).json({
      ok: true,
      count: operations.length,
      operations,
    });
  } catch (err) {
    return res.status(500).json({
      ok: false,
      error: err.message,
    });
  }
}
