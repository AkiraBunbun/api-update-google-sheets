api/update.js:
```javascript
import { GoogleSpreadsheet } from 'google-spreadsheet';

export default async function(req, res) {
  const { rango, valores } = req.body;
  const doc = new GoogleSpreadsheet(process.env.SHEET_ID);
  await doc.useServiceAccountAuth({ client_email: process.env.CLIENT_EMAIL, private_key: process.env.PRIVATE_KEY });
  await doc.loadInfo();
  const sheet = doc.sheetsByIndex[0];
  await sheet.updateCells(sheet.getRange(rango).cellList);  // Ultra-rápido
  res.json({ ok: true, ms: Date.now() - req.start });
}
