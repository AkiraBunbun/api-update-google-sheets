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

  try {
    const doc = new GoogleSpreadsheet(process.env.SHEET_ID);

    // Use the private key directly from the environment variable
    const private_key = process.env.PRIVATE_KEY;  // Directly use the value from environment variable

    const auth = new GoogleAuth({
      credentials: {
        client_email: process.env.CLIENT_EMAIL,
        private_key: private_key.replace(/\\n/g, '\n')  // Ensure proper newline decoding
      },
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const authClient = await auth.getClient();
    doc.useOAuth2Client(authClient);

    await doc.loadInfo();  // Load the document information
    const sheet = doc.sheetsByIndex[0];  // Access the first sheet

    // Get all rows from the sheet
    const rows = await sheet.getRows();

    // Find the row and column index based on the range `rango` (e.g., 'D1')
    const [column, row] = rango.match(/[A-Z]+|[0-9]+/g); // Extract column (e.g., 'D') and row (e.g., '1')

    // Adjust the column index based on the column letter (A -> 0, B -> 1, C -> 2, ...)
    const columnIndex = column.charCodeAt(0) - 65;  // Assuming column A starts at 0, B at 1, ...

    // Update the correct cell value
    rows[row - 1][sheet.headerValues[columnIndex]] = valores[0][0];  // Set the value of the cell

    // Save the row to the sheet
    await rows[row - 1].save();

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
