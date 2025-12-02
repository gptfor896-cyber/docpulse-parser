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

    // Попробуем найти лист, в названии которого есть "Отчет" (если нет — берём первый)
    const targetSheetName =
      workbook.SheetNames.find((n) =>
        String(n).toLowerCase().includes("отчет")
      ) || workbook.SheetNames[0];

    const sheet = workbook.Sheets[targetSheetName];

    if (!sheet) {
      return res.status(400).json({
        ok: false,
        error: "Target sheet not found",
        sheetNames: workbook.SheetNames,
      });
    }

    // Преобразуем лист в массив строк
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1, // [ [ячейки 1 строки], [ячейки 2 строки] ... ]
      raw: true,
    });

    if (!rows || rows.length === 0) {
      return res.status(400).json({
        ok: false,
        error: "Sheet is empty",
      });
    }

    // 4. Ищем строку с заголовками — там, где встречается "Артикул"
    let headerRowIndex = -1;

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;

      const joined = row
        .map((c) => (c === null || c === undefined ? "" : String(c)))
        .join(" ")
        .toLowerCase();

      if (joined.includes("артикул продавца") || joined.includes("артикул")) {
        headerRowIndex = i;
        break;
      }
    }

    if (headerRowIndex === -1) {
      // Не нашли вообще строку с "Артикул" — вернём немного дебаг-инфы
      return res.status(400).json({
        ok: false,
        error: "Header row with 'Артикул' not found in any row",
        sheetNames: workbook.SheetNames,
        sampleFirstRows: rows.slice(0, 25),
      });
    }

    const headerRow = rows[headerRowIndex] || [];

    const headers = headerRow.map((h) =>
      h === undefined || h === null ? "" : String(h).trim()
    );

    const col = (name) => headers.indexOf(name);

    const colSku = col("Артикул продавца");
    const colQtySale = col("Количество");
    const colAmountSale = col("Итого к начислению, руб.");
    const colQtyReturn = col("Количество возвратов");
    const colAmountReturn = col("Итого возвращено, руб.");

    if (colSku === -1) {
      // Не нашли нужное имя колонки даже в найденной строке заголовка
      return res.status(400).json({
        ok: false,
        error:
          "Column 'Артикул продавца' not found in detected header row",
        detectedHeaderRowIndex: headerRowIndex + 1, // человеко-читаемый (1-based)
        headers,
      });
    }

    const operations = [];

    // 5. Проходим по всем строкам ниже заголовков
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
      const rawAmountSale =
        colAmountSale > -1 ? row[colAmountSale] ?? 0 : 0;

      const rawQtyReturn =
        colQtyReturn > -1 ? row[colQtyReturn] ?? 0 : 0;
      const rawAmountReturn =
        colAmountReturn > -1 ? row[colAmountReturn] ?? 0 : 0;

      const qtySale = Number(rawQtySale) || 0;
      const amountSale =
        Number(
          typeof rawAmountSale === "string"
            ? rawAmountSale.replace(",", ".")
            : rawAmountSale
        ) || 0;

      const qtyReturn = Number(rawQtyReturn) || 0;
      const amountReturn =
        Number(
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
