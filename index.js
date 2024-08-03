const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const multer = require('multer');
const { google } = require('googleapis');
const { sheetValuesToObject } = require('./utils');
const { config } = require('dotenv');
const { jwtClient } = require('./google');
config();

const app = express();
const router = express.Router();
const PORT = process.env.PORT || 3001;

app.use(bodyParser.json());
app.use(cors());

// Ruta para la raíz
app.get('/', (req, res) => {
  res.send('El servidor está funcionando correctamente');
});

// Ruta para obtener datos de Google Sheets
router.post('/getData', async (req, res) => {
  try {
    const { sheetName } = req.body;
    const spreadsheetId = '1sp9G8A6-hPUtnmfK7jSpAqQfoMzKR3kmYGYzhOAC6vM';
    const range = `${sheetName}!A1:Z1000`; // Rango a obtener
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
    });

    const values = response.data.values;
    const data = sheetValuesToObject(values);

    res.status(200).json({ status: true, data });
  } catch (error) {
    console.error('Error:', error);
    res.status(400).json({ status: false, error });
  }
});

// Ruta para actualizar datos en Google Sheets
router.post('/updateData', async (req, res) => {
  try {
    const { updateData, id, sheetName } = req.body;
    const spreadsheetId = '1sp9G8A6-hPUtnmfK7jSpAqQfoMzKR3kmYGYzhOAC6vM';
    const range = `${sheetName}!A1:Z1000`;
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });

    const responseSheet = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
    });

    const currentValues = responseSheet.data.values;
    const rowIndex = currentValues.findIndex(row => row[0] == id);

    if (rowIndex === -1) {
      return res.status(404).json({ error: 'ID no encontrado', status: false });
    }

    const updatedRange = `${sheetName}!A${rowIndex + 1}`;
    const sheetsResponse = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: updatedRange,
      valueInputOption: 'RAW',
      resource: {
        values: [updateData],
      },
    });

    if (sheetsResponse.status === 200) {
      return res.status(200).json({ success: 'Se actualizó correctamente', status: true });
    } else {
      return res.status(400).json({ error: 'No se actualizó', status: false });
    }
  } catch (error) {
    console.error('Error en la conexión:', error);
    return res.status(400).json({ error: 'Error en la conexión', status: false });
  }
});

// Ruta para actualizar metas y avances en Google Sheets
router.post('/updateMetas', async (req, res) => {
  try {
    const { updateData, id, sheetName } = req.body;
    const spreadsheetId = '1sp9G8A6-hPUtnmfK7jSpAqQfoMzKR3kmYGYzhOAC6vM';
    const range = `${sheetName}!A1:Z1000`;
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });

    const responseSheet = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
    });

    const currentValues = responseSheet.data.values;
    const rowIndex = currentValues.findIndex(row => row[0] == id);

    if (rowIndex === -1) {
      return res.status(404).json({ error: 'ID no encontrado', status: false });
    }

    const updatedRange = `${sheetName}!A${rowIndex + 1}:${rowIndex + 1}`;
    const updatedValues = [...currentValues[rowIndex]];

    // Buscar el índice de la columna que queremos actualizar
    const columnIndex = currentValues[0].indexOf(updateData[0]);
    if (columnIndex === -1) {
      return res.status(400).json({ error: 'Columna no encontrada', status: false });
    }

    // Actualizar el valor en la columna correspondiente
    updatedValues[columnIndex] = updateData[1];

    // Recalcular meta_trienio y total_ejec
    const meta2024 = parseFloat(updatedValues[currentValues[0].indexOf('2024')]) || 0;
    const meta2025 = parseFloat(updatedValues[currentValues[0].indexOf('2025')]) || 0;
    const meta2026 = parseFloat(updatedValues[currentValues[0].indexOf('2026')]) || 0;
    const ejec2024 = parseFloat(updatedValues[currentValues[0].indexOf('ejec_2024')]) || 0;
    const ejec2025 = parseFloat(updatedValues[currentValues[0].indexOf('ejec_2025')]) || 0;
    const ejec2026 = parseFloat(updatedValues[currentValues[0].indexOf('ejec_2026')]) || 0;

    updatedValues[currentValues[0].indexOf('meta_trienio')] = meta2024 + meta2025 + meta2026;
    updatedValues[currentValues[0].indexOf('total_ejec')] = ejec2024 + ejec2025 + ejec2026;

    const sheetsResponse = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: updatedRange,
      valueInputOption: 'RAW',
      resource: {
        values: [updatedValues],
      },
    });

    if (sheetsResponse.status === 200) {
      return res.status(200).json({ success: 'Se actualizó correctamente', status: true });
    } else {
      return res.status(400).json({ error: 'No se actualizó', status: false });
    }
  } catch (error) {
    console.error('Error en la conexión:', error);
    return res.status(400).json({ error: 'Error en la conexión', status: false });
  }
});


app.use(router);

app.listen(PORT, () => {
  console.log(`Servidor backend escuchando en el puerto ${PORT}`);
});
