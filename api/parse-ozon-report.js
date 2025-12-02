// Простейший обработчик для Vercel
export default async function handler(req, res) {
  // Временно разрешаем любые методы, чтобы было проще тестировать
  res.status(200).json({
    ok: true,
    message: 'DocPulse XLSX parser is alive',
    method: req.method
  });
}
