const { GoogleSpreadsheet } = require('google-spreadsheet');
const { GoogleAuth } = require('google-auth-library');

module.exports = async function handler(req, res) {
  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  // Handle preflight OPTIONS request
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') return res.status(405).json({ error: 'Solo POST' });

  const inicio = Date.now();
  const { rango = 'main!A1', valores } = req.body;

  try {
    const doc = new GoogleSpreadsheet(process.env.SHEET_ID);

    // Ensure the private key is correctly formatted
    const private_key = `-----BEGIN PRIVATE KEY-----
    ${process.env.PRIVATE_KEY}
    -----END PRIVATE KEY-----`;

    const auth = new GoogleAuth({
      credentials: {
        client_email: process.env.CLIENT_EMAIL,
        private_key: private_key.replace(/\\n/g, '\n')
      },
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const authClient = await auth.getClient();
    doc.useOAuth2Client(authClient);

    await doc.loadInfo();
    const sheet = doc.sheetsByIndex[0];  // "main" sheet

    const range = sheet.getRange(rango);
    const cells = range._cellCells || [];  // Reuse if exists
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
    console.error('Error:', error.message);
    res.status(500).json({ error: error.message });
  }
};
