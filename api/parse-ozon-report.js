export const config = {
  runtime: "nodejs",
};

import * as XLSX from "xlsx";

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ ok: false, error: "Use POST" });
    }

    const { fileUrl } = req.body || {};
    if (!fileUrl) {
      return res.status(400).json({ ok: false, error: "Missing fileUrl" });
    }

    // 1. Скачиваем файл
    const response = await fetch(fileUrl);
    if (!response.ok) {
      return res.status(400).json({
        ok: false,
        error: `Failed to download file: ${response.status}`,
      });
    }

    const arrayBuffer = await response.arrayBuffer();
    const uint8 = new Uint8Array(arrayBuffer);

    // 2. Читаем XLSX через SheetJS
    const workbook = XLSX.read(uint8, { type: "array" });

    const sheetName =
      workbook.SheetNames.find((n) =>
        String(n).toLowerCase().includes("отчет")
      ) || workbook.SheetNames[0];

    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      return res.status(400).json({
        ok: false,
        error: "Target sheet not found",
        sheetNames: workbook.SheetNames,
      });
    }

    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: true,
    });

    if (!rows || rows.length === 0) {
      return res.status(400).json({ ok: false, error: "Sheet is empty" });
    }

    // 3. Находим первую строку данных: где в первом столбце число (1, 2, 3...)
    let dataStartIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const cell0 = row[0];
      if (typeof cell0 === "number") {
        dataStartIndex = i;
        break;
      }
    }

    if (dataStartIndex === -1) {
      return res.status(400).json({
        ok: false,
        error: "No data rows found (no numeric index in first column)",
      });
    }

    const getNum = (val) => {
      if (val === null || val === undefined || val === "") return 0;
      if (typeof val === "number") return val;
      const s = String(val).replace(" ", "").replace(",", ".");
      const parsed = Number(s);
      return isNaN(parsed) ? 0 : parsed;
    };

    const operations = [];

    // 4. Проходим по всем строкам данных
    for (let i = dataStartIndex; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;

      // Таблица кончилась, когда первый столбец перестал быть числом
      const rowNum = row[0];
      if (typeof rowNum !== "number") break;

      // Могут быть обрезанные строки, поэтому берём через || null
      const sku = row[4] ?? null;          // колонка 5 (SKU)
      if (!sku) continue;

      const qtySale = getNum(row[8]);      // колонка 9 — Кол-во продаж
      const amountSale = getNum(row[13]);  // колонка 14 — Итого к начислению, руб.

      const qtyReturn = getNum(row[17]);   // колонка 18 — Кол-во возвратов
      const amountReturn = getNum(row[20]); // колонка 21 — Итого возвращено, руб.

      const orderNumber = row[21] ?? null; // колонка 22 — Номер отправления
      const orderDateVal = row[22] ?? null; // колонка 23 — Дата отправления

      let orderDate = null;
      if (orderDateVal instanceof Date) {
        orderDate = orderDateVal.toISOString().slice(0, 10);
      } else if (typeof orderDateVal === "string") {
        orderDate = orderDateVal;
      }

      // Продажа
      if (amountSale !== 0) {
        operations.push({
          operation_type: "sale",
          sku: String(sku),
          quantity: qtySale,
          amount: amountSale,
          order_number: orderNumber ? String(orderNumber) : null,
          order_date: orderDate,
        });
      }

      // Возврат
      if (amountReturn !== 0) {
        operations.push({
          operation_type: "return",
          sku: String(sku),
          quantity: qtyReturn,
          amount: -Math.abs(amountReturn),
          order_number: orderNumber ? String(orderNumber) : null,
          order_date: orderDate,
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
