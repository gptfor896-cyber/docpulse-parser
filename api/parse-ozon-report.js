import ExcelJS from "exceljs";
import fetch from "node-fetch";

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

    // 1. Скачиваем файл
    const response = await fetch(fileUrl);
    const buffer = await response.arrayBuffer();

    // 2. Читаем Excel
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(Buffer.from(buffer));

    const sheet = workbook.getWorksheet(1); // Первый лист

    let operations = [];

    // Найти строку заголовков (Ищем ячейку "Артикул продавца")
    let headerRowIndex = null;

    sheet.eachRow((row, rowNumber) => {
      const text = row.values.join(" ");
      if (text.includes("Артикул продавца")) {
        headerRowIndex = rowNumber;
      }
    });

    if (!headerRowIndex) {
      return res.status(400).json({
        ok: false,
        error: "Header row not found",
      });
    }

    // 3. Читаем строки ниже заголовков
    const headerRow = sheet.getRow(headerRowIndex);
    const headers = headerRow.values.map((h) =>
      typeof h === "string" ? h.trim() : h
    );

    const col = (name) => headers.indexOf(name);

    for (let i = headerRowIndex + 1; i <= sheet.rowCount; i++) {
      const row = sheet.getRow(i);
      if (!row || !row.values || row.values.length < 5) continue;

      const sku = row.getCell(col("Артикул продавца")).value;
      if (!sku) continue;

      const qtySale = row.getCell(col("Количество")).value || 0;
      const amountSale = row.getCell(col("Итого к начислению, руб.")).value || 0;

      const qtyReturn =
        row.getCell(col("Количество возвратов")).value || 0;
      const amountReturn =
        row.getCell(col("Итого возвращено, руб.")).value || 0;

      // Продажа
      if (amountSale > 0) {
        operations.push({
          operation_type: "sale",
          sku,
          quantity: qtySale,
          amount: amountSale,
        });
      }

      // Возврат
      if (amountReturn > 0) {
        operations.push({
          operation_type: "return",
          sku,
          quantity: qtyReturn,
          amount: -Math.abs(amountReturn),
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
