import { GoogleSpreadsheet } from 'google-spreadsheet';

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Solo POST' });

  const inicio = Date.now();
  const { rango = 'main!A1', valores } = req.body;

  try {
    const doc = new GoogleSpreadsheet(process.env.SHEET_ID);
    await doc.useServiceAccountAuth({
      client_email: process.env.CLIENT_EMAIL,
      private_key: process.env.PRIVATE_KEY.replace(/\\n/g, '\n')  // ← Maneja saltos
    });
    await doc.loadInfo();
    const sheet = doc.sheetsByIndex[0];  // Hoja "main"
    
    // **Update ultra-rápido**
    const range = sheet.getRange(rango);
    const cells = range._cellCells || [];  // Reutiliza si existe
    await sheet.updateCells(cells.map((cell, i) => ({
      ...cell,
      value: valores.flat()[i] || ''
    })));

    res.json({
      ok: true,
      ms: Date.now() - inicio,
      updated: valores
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
}
