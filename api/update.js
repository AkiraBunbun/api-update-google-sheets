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
  const { rango, valores } = req.body;

  // Input validation
  if (!rango || !valores || !Array.isArray(valores)) {
    return res.status(400).json({ error: 'Invalid input data' });
  }

  try {
    const doc = new GoogleSpreadsheet(process.env.SHEET_ID);

    // Use the private key directly from the environment variable
    const private_key = process.env.PRIVATE_KEY.replace(/\\n/g, '\n'); // Ensure newlines are correct

    const auth = new GoogleAuth({
      credentials: {
        client_email: process.env.CLIENT_EMAIL,
        private_key,
      },
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const authClient = await auth.getClient();
    doc.useOAuth2Client(authClient);

    await doc.loadInfo(); // Load document info
    const sheet = doc.sheetsByIndex[0]; // First sheet

    const rows = await sheet.getRows(); // Get all rows

    // Parse the range (e.g., 'D1')
    const [colLetter, rowNumber] = rango.match(/[A-Z]+|[0-9]+/g);
    const columnIndex = colLetter.charCodeAt(0) - 65; // Column A = 0, B = 1, C = 2, etc.

    // Make sure the row exists
    const rowIndex = parseInt(rowNumber, 10) - 1;
    if (rowIndex < 0 || rowIndex >= rows.length) {
      return res.status(400).json({ error: 'Row index out of range' });
    }

    const header = sheet.headerValues; // Get column headers
    const columnHeader = header[columnIndex]; // Get column header name

    // Log for debugging
    console.log('Header:', header);
    console.log('Updating row:', rowIndex, 'Column:', columnHeader, 'Value:', valores[0][0]);

    // Update the cell value
    rows[rowIndex][columnHeader] = valores[0][0]; // Update the value based on header name

    // Save the row
    await rows[rowIndex].save();

    res.json({
      ok: true,
      ms: Date.now() - inicio,
      updated: valores,
    });
  } catch (error) {
    console.error('Error:', error.message);
    res.status(500).json({ error: error.message });
  }
};
