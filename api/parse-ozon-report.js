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
        error: "Sheet not found",
        sheetNames: workbook.SheetNames,
      });
    }

    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: true,
    });

    if (!rows || !rows.length) {
      return res.status(400).json({ ok: false, error: "Sheet is empty" });
    }

    // 3. Находим первую строку данных – где первый столбец число (1, 2, 3...)
    let dataStartIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const v = row[0];
      if (typeof v === "number") {
        dataStartIndex = i;
        break;
      }
    }

    if (dataStartIndex === -1) {
      return res.status(400).json({
        ok: false,
        error: "No data rows (no numeric index in first column)",
        sampleFirstRows: rows.slice(0, 25),
      });
    }

    // Индексы колонок по позиции (0-based)
    const IDX_SKU = 4;            // E (5-й столбец)
    const IDX_QTY_SALE = 8;       // I (9-й)
    const IDX_AMOUNT_SALE = 13;   // N (14-й, 'Итого к начислению, руб.')
    const IDX_QTY_RETURN = 17;    // R (18-й)
    const IDX_AMOUNT_RETURN = 20; // U (21-й, 'Итого возвращено, руб.')
    const IDX_ORDER_NUMBER = 21;  // V (22-й)
    const IDX_ORDER_DATE = 22;    // W (23-й)

    const toNumber = (val) => {
      if (val === null || val === undefined || val === "") return 0;
      if (typeof val === "number") return val;
      const s = String(val).replace(/\s/g, "").replace(",", ".");
      const n = Number(s);
      return isNaN(n) ? 0 : n;
    };

    // Преобразование excel-даты (кол-во дней с 1899-12-30) в YYYY-MM-DD
    const excelDateToISO = (num) => {
      if (typeof num !== "number") return null;
      const excelEpoch = new Date(Date.UTC(1899, 11, 30)); // 1899-12-30
      const msPerDay = 24 * 60 * 60 * 1000;
      const d = new Date(excelEpoch.getTime() + num * msPerDay);
      const year = d.getUTCFullYear();
      const month = String(d.getUTCMonth() + 1).padStart(2, "0");
      const day = String(d.getUTCDate()).padStart(2, "0");
      return `${year}-${month}-${day}`;
    };

    const operations = [];

    // 4. Проходим по строкам данных
    for (let i = dataStartIndex; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;

      const idxVal = row[0];
      if (typeof idxVal !== "number") break; // таблица кончилась

      const sku = row[IDX_SKU];
      if (!sku) continue;

      const qtySale = toNumber(row[IDX_QTY_SALE]);
      const amountSale = toNumber(row[IDX_AMOUNT_SALE]);

      const qtyReturn = toNumber(row[IDX_QTY_RETURN]);
      const amountReturn = toNumber(row[IDX_AMOUNT_RETURN]);

      const orderNumber = row[IDX_ORDER_NUMBER] ?? null;
      const orderDateVal = row[IDX_ORDER_DATE] ?? null;

      let orderDate = null;

      if (orderDateVal instanceof Date) {
        orderDate = orderDateVal.toISOString().slice(0, 10);
      } else if (typeof orderDateVal === "number") {
        // excel serial date
        orderDate = excelDateToISO(orderDateVal);
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
