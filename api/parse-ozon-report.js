export const config = {
  runtime: "nodejs", // работаем в обычном Node.js окружении
};

import ExcelJS from "exceljs";

export default async function handler(req, res) {
  try {
    // 1. Проверяем метод
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

    // 2. Скачиваем файл (используем встроенный fetch, БЕЗ node-fetch)
    const response = await fetch(fileUrl);
    if (!response.ok) {
      return res.status(400).json({
        ok: false,
        error: `Failed to download file: ${response.status}`,
      });
    }

    const arrayBuffer = await response.arrayBuffer();

    // 3. Читаем Excel
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(Buffer.from(arrayBuffer));

    // Берём первый лист
    const sheet = workbook.getWorksheet(1);
    if (!sheet) {
      return res.status(400).json({
        ok: false,
        error: "No worksheet found in workbook",
      });
    }

    // 🧠 ВАЖНО: фиксируем номер строки с заголовками
    const headerRowIndex = 14; // ты говорил: на 14 строке заголовки
    const headerRow = sheet.getRow(headerRowIndex);

    // Получаем массив заголовков
    const headers = headerRow.values.map((h) =>
      typeof h === "string" ? h.trim() : ""
    );

    // Удобная функция: по имени колонки получить её индекс
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

    let operations = [];

    // 4. Идём по всем строкам НИЖЕ заголовков
    for (let i = headerRowIndex + 1; i <= sheet.rowCount; i++) {
      const row = sheet.getRow(i);
      if (!row || !row.values) continue;

      // Если строка совсем пустая — пропускаем
      const isEmpty = row.values
        .slice(1)
        .every((v) => v === null || v === undefined || v === "");
      if (isEmpty) continue;

      const sku = colSku > -1 ? row.getCell(colSku).value : null;
      if (!sku) continue; // строка без артикула нам не нужна

      const rawQtySale =
        colQtySale > -1 ? row.getCell(colQtySale).value ?? 0 : 0;
      const rawAmountSale =
        colAmountSale > -1 ? row.getCell(colAmountSale).value ?? 0 : 0;

      const rawQtyReturn =
        colQtyReturn > -1 ? row.getCell(colQtyReturn).value ?? 0 : 0;
      const rawAmountReturn =
        colAmountReturn > -1 ? row.getCell(colAmountReturn).value ?? 0 : 0;

      const qtySale = Number(rawQtySale) || 0;
      const amountSale = Number(rawAmountSale) || 0;

      const qtyReturn = Number(rawQtyReturn) || 0;
      const amountReturn = Number(rawAmountReturn) || 0;

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
    stack: err.stack,   // 👈 добавили
  });
}

}
