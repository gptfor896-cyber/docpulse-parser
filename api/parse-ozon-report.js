// Простейший обработчик для Vercel
export default async function handler(req, res) {
  // Пока разрешаем только POST — дальше сюда будем слать URL файла
  if (req.method !== 'POST') {
    res.status(405).json({ ok: false, error: 'Use POST' });
    return;
  }

  // Заглушка: просто проверка, что функция живая
  res.status(200).json({
    ok: true,
    message: 'DocPulse XLSX parser is alive'
  });
}

