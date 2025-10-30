const { GoogleSpreadsheet } = require('google-spreadsheet');
const { GoogleAuth } = require('google-auth-library');

module.exports = async function handler(req, res) {
  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*'); // Allow requests from any origin (replace '*' with your frontend domain for tighter security)
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS'); // Allow specific HTTP methods
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization'); // Allow specific headers

  // Handle preflight OPTIONS request
  if (req.method === 'OPTIONS') {
    return res.status(200).end(); // Respond successfully to preflight request
  }

  if (req.method !== 'POST') return res.status(405).json({ error: 'Solo POST' });

  const inicio = Date.now();
  const { rango = 'main!A1', valores } = req.body;

  try {
    const doc = new GoogleSpreadsheet(process.env.SHEET_ID);

    // Create a GoogleAuth instance for service account authentication
    const auth = new GoogleAuth({
      credentials: {
        client_email: process.env.CLIENT_EMAIL,
        private_key: process.env.PRIVATE_KEY.replace(/\\n/g, '\n')
      },
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    // Use the GoogleAuth client to authenticate
    const authClient = await auth.getClient();
    doc.useOAuth2Client(authClient);  // Use the OAuth2 client for auth

    await doc.loadInfo();
    const sheet = doc.sheetsByIndex[0];  // "main" sheet

    // **Update ultra-fast**
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
    res.status(500).json({ error: error.message });
  }
};
