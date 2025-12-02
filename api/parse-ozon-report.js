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
      return res
        .status(400)
        .json({ ok: false, error: `Failed to download file: ${response.status}` });
    }

    const arrayBuffer = await response.arrayBuffer();
    const uint8 = new Uint8Array(arrayBuffer);

    // 2. Читаем XLSX
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

    // 3. Находим строку с "№ п/п" — верхняя шапка (13-я в твоём файле)
    let headerTopIndex = -1;

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const first = row[0];
      if (typeof first === "string" && first.trim().startsWith("№ п/п")) {
        headerTopIndex = i;
        break;
      }
    }

    if (headerTopIndex === -1) {
      return res.status(400).json({
        ok: false,
        error: "Top header row with '№ п/п' not found",
        sampleFirstRows: rows.slice(0, 25),
      });
    }

    const headerTop = rows[headerTopIndex] || [];
    const headerSecond = rows[headerTopIndex + 1] || [];

    // Нормализуем строки
    const normRow = (row) =>
      row.map((v) =>
        v === null || v === undefined ? "" : String(v).trim()
      );

    const top = normRow(headerTop);
    const second = normRow(headerSecond);

    // Функции поиска индексов
    const findIndex = (row, predicate) =>
      row.findIndex((v, idx) => predicate(v, idx));

    const topIdx = (name) =>
      findIndex(top, (v) => v === name || v.includes(name));
    const secondIdx = (name) =>
      findIndex(second, (v) => v === name || v.includes(name));

    // 4. Находим нужные колонки

    // SKU (из первой шапки)
    const colSku = topIdx("SKU"); // в твоём файле это 4-й столбец
    if (colSku === -1) {
      return res.status(400).json({
        ok: false,
        error: "Column 'SKU' not found in top header",
        headerTop: top,
      });
    }

    // Кол-во продаж и возвратов (второй ряд шапки)
    const qtyCols = [];
    second.forEach((v, idx) => {
      if (v === "Кол-во") qtyCols.push(idx);
    });

    const colQtySale = qtyCols[0] ?? -1;   // под блоком "Реализовано"
    const colQtyReturn = qtyCols[1] ?? -1; // под блоком "Возвращено клиентом"

    // Сумма продажи и сумма возврата
    const colAmountSale = secondIdx("Итого к начислению, руб.");
    const colAmountReturn = secondIdx("Итого возвращено, руб.");

    // Номер и дата отправления (последний блок "Отправление")
    const colOrderNumber = secondIdx("Номер"); // но таких может быть 2 — нам нужен тот, что ближе к концу
    const colOrderDate = secondIdx("Дата");    // тут тоже 2 "Дата", ниже уточним

    // Уточняем номер и дату отправления: берём те столбцы, где top === "Отправление"
    const colBlockOtpravlenie = topIdx("Отправление");
    let realOrderNumberCol = -1;
    let realOrderDateCol = -1;

    for (let idx = 0; idx < second.length; idx++) {
      if (second[idx] === "Номер" && idx >= colBlockOtpravlenie) {
        realOrderNumberCol = idx;
      }
      if (second[idx] === "Дата" && idx >= colBlockOtpravlenie) {
        // первый "Дата" после блока "Отправление" считаем датой отправления
        if (realOrderDateCol === -1) realOrderDateCol = idx;
      }
    }

    // 5. Находим первую строку данных — где первый столбец = 1, 2, 3 и т.д.
    let dataStartIndex = -1;
    for (let i = headerTopIndex + 2; i < rows.length; i++) {
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
        error: "No data rows found after header",
      });
    }

    const operations = [];

    // 6. Проходим по строкам данных
    for (let i = dataStartIndex; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;

      // Проверяем, не пустая ли строка
      const isEmpty = row.every(
        (v) => v === null || v === undefined || v === ""
      );
      if (isEmpty) continue;

      const rowNum = row[0];
      if (typeof rowNum !== "number") {
        // Если первый столбец перестал быть числом — значит, таблица кончилась
        break;
      }

      const sku = row[colSku];
      if (!sku) continue;

      const getNum = (val) => {
        if (val === null || val === undefined || val === "") return 0;
        if (typeof val === "number") return val;
        const s = String(val).replace(" ", "").replace(",", ".");
        const parsed = Number(s);
        return isNaN(parsed) ? 0 : parsed;
      };

      const qtySale = colQtySale !== -1 ? getNum(row[colQtySale]) : 0;
      const amountSale =
        colAmountSale !== -1 ? getNum(row[colAmountSale]) : 0;

      const qtyReturn =
        colQtyReturn !== -1 ? getNum(row[colQtyReturn]) : 0;
      const amountReturn =
        colAmountReturn !== -1 ? getNum(row[colAmountReturn]) : 0;

      const orderNumber =
        realOrderNumberCol !== -1 ? row[realOrderNumberCol] : null;
      const orderDateVal =
        realOrderDateCol !== -1 ? row[realOrderDateCol] : null;

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
