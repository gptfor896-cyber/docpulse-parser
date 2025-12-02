export const config = {
  runtime: "nodejs",
};

import * as XLSX from "xlsx";

export default async function handler(req, res) {
  try {
    // 1. Только POST
    if (req.method !== "POST") {
      return res.status(405).json({
        ok: false,
        error: "Use POST",
      });
    }

    const { fileUrl } = req.body || {};

    if (!fileUrl) {
      return res.status(400).json({
        ok: false,
        error: "Missing fileUrl",
      });
    }

    // 2. Скачиваем файл
    const response = await fetch(fileUrl);
    if (!response.ok) {
      return res.status(400).json({
        ok: false,
        error: `Failed to download file: ${response.status}`,
      });
    }

    const arrayBuffer = await response.arrayBuffer();
    const uint8 = new Uint8Array(arrayBuffer);

    // 3. Читаем XLSX через SheetJS
    const workbook = XLSX.read(uint8, { type: "array" });

    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
      return res.status(400).json({
        ok: false,
        error: "No sheets in workbook",
      });
    }

    const sheet = workbook.Sheets[firstSheetName];

    // Преобразуем лист в массив строк (каждая строка — массив ячеек)
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,    // вернёт [ [ячейки первой строки], [ячейки второй] ... ]
      raw: true,
    });

    if (!rows || rows.length === 0) {
      return res.status(400).json({
        ok: false,
        error: "Sheet is empty",
      });
    }

    // 🧠 ВАЖНО: в Excel заголовки на 14-й строке → индекс 13 (0-based)
    const headerRowIndex = 13;
    const headerRow = rows[headerRowIndex] || [];

    const headers = headerRow.map((h) =>
      h === undefined || h === null ? "" : String(h).trim()
    );

    // Функция: получить индекс колонки по названию
    const col = (name) => headers.indexOf(name);

    const colSku = col("Артикул продавца");
    const colQtySale = col("Количество");
    const colAmountSale = col("Итого к начислению, руб.");
    const colQtyReturn = col("Количество возвратов");
    const colAmountReturn = col("Итого возвращено, руб.");

    if (colSku === -1) {
      return res.status(400).json({
        ok: false,
        error: "Column 'Артикул продавца' not found in header row 14",
        headers,
      });
    }

    const operations = [];

    // 4. Проходим по всем строкам ниже заголовков
    for (let i = headerRowIndex + 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;

      // Проверяем, пустая ли строка
      const isEmpty = row.every(
        (v) => v === null || v === undefined || v === ""
      );
      if (isEmpty) continue;

      const sku = colSku > -1 ? row[colSku] : null;
      if (!sku) continue;

      const rawQtySale = colQtySale > -1 ? row[colQtySale] ?? 0 : 0;
      const rawAmountSale = colAmountSale > -1 ? row[colAmountSale] ?? 0 : 0;

      const rawQtyReturn =
        colQtyReturn > -1 ? row[colQtyReturn] ?? 0 : 0;
      const rawAmountReturn =
        colAmountReturn > -1 ? row[colAmountReturn] ?? 0 : 0;

      const qtySale = Number(rawQtySale) || 0;
      const amountSale = Number(
        typeof rawAmountSale === "string"
          ? rawAmountSale.replace(",", ".")
          : rawAmountSale
      ) || 0;

      const qtyReturn = Number(rawQtyReturn) || 0;
      const amountReturn = Number(
        typeof rawAmountReturn === "string"
          ? rawAmountReturn.replace(",", ".")
          : rawAmountReturn
      ) || 0;

      // Продажа
      if (amountSale !== 0) {
        operations.push({
          operation_type: "sale",
          sku: String(sku),
          quantity: qtySale,
          amount: amountSale,
        });
      }

      // Возврат
      if (amountReturn !== 0) {
        operations.push({
          operation_type: "return",
          sku: String(sku),
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
      stack: err.stack,
    });
  }
}
