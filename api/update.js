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

  const { sheetId, sheetName, object, requestType } = req.body;

  if (!sheetId || !sheetName || !object || !requestType) {
    return res.status(400).json({ error: 'Missing sheetId, sheetName, object, or requestType in the request body.' });
  }

  try {
    // Initialize Google Spreadsheet API
    const doc = new GoogleSpreadsheet(sheetId);
    const private_key = process.env.PRIVATE_KEY.replace(/\\n/g, '\n');  // Ensure correct newlines

    const auth = new GoogleAuth({
      credentials: {
        client_email: process.env.CLIENT_EMAIL,
        private_key,
      },
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const authClient = await auth.getClient();
    doc.useOAuth2Client(authClient);

    await doc.loadInfo();  // Load document info
    const sheet = doc.sheetsByTitle[sheetName];  // Access the sheet by name

    if (!sheet) {
      return res.status(404).json({ error: `Sheet with name ${sheetName} not found.` });
    }

    const rows = await sheet.getRows(); // Get all rows

    // Process based on requestType
    switch (requestType) {
      case 'update':
        return await updateRow(rows, sheet, object, res);
        
      case 'create':
        return await createRow(sheet, object, res);

      case 'delete':
        return await deleteRow(rows, sheet, object, res);

      default:
        return res.status(400).json({ error: 'Invalid request type' });
    }
  } catch (error) {
    console.error('Error:', error.message);
    return res.status(500).json({ error: error.message });
  }
};

// Function to update a row
async function updateRow(rows, sheet, object, res) {
  const rowIndex = rows.findIndex(row => row.id === object.id.toString()); // Find row by ID

  if (rowIndex >= 0) {
    const row = rows[rowIndex];
    for (const [key, value] of Object.entries(object)) {
      if (key !== 'id' && row[key] !== undefined) {
        row[key] = value; // Update the field if it exists in the row
      }
    }
    await row.save();  // Save the updated row

    // Return only the fields we need, avoiding the circular structure
    const updatedRow = {
      id: row.id,
      name: row.name,
      phone: row.phone,
      // Include other fields as needed
    };

    return res.status(200).json({ message: 'Row updated successfully.', updatedRow });
  } else {
    return res.status(404).json({ error: `Row with ID ${object.id} not found.` });
  }
}

// Function to create a new row
async function createRow(sheet, object, res) {
  if (!object.id) {
    return res.status(400).json({ error: 'Missing ID for new row creation.' });
  }

  const newRow = await sheet.addRow(object);  // Add the new row

  // Return only the fields we need, avoiding the circular structure
  const createdRow = {
    id: newRow.id,
    name: newRow.name,
    phone: newRow.phone,
    // Include other fields as needed
  };

  return res.status(201).json({ message: 'New row created successfully.', createdRow });
}

// Function to delete a row
async function deleteRow(rows, sheet, object, res) {
  const rowToDelete = rows.find(row => row.id === object.id.toString());

  if (rowToDelete) {
    await rowToDelete.delete();  // Delete the row if found
    return res.status(200).json({ message: 'Row deleted successfully.' });
  } else {
    return res.status(404).json({ error: `Row with ID ${object.id} not found.` });
  }
}
