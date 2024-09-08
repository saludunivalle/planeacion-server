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

// Ruta raíz para verificar si el servidor está funcionando
app.get('/', (req, res) => {
  res.send('El servidor está funcionando correctamente');
});

// Ruta para obtener datos de Google Sheets
app.post('/getData', async (req, res) => {
  try {
    const { sheetName } = req.body;
    const spreadsheetId = '1sp9G8A6-hPUtnmfK7jSpAqQfoMzKR3kmYGYzhOAC6vM'; 
    const range = `${sheetName}!A1:Z1000`; 
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
    });

    const values = response.data.values;
    const data = sheetValuesToObject(values); 

    res.status(200).json({ status: true, data });
  } catch (error) {
    console.error('Error obteniendo datos de Google Sheets:', error);
    res.status(400).json({ status: false, error });
  }
});

// Ruta para actualizar datos en Google Sheets
app.post('/updateData', async (req, res) => {
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
    console.error('Error actualizando datos en Google Sheets:', error);
    return res.status(400).json({ error: 'Error en la conexión', status: false });
  }
});

// Ruta para actualizar metas en Google Sheets
app.post('/updateMetas', async (req, res) => {
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
    console.error('Error actualizando metas en Google Sheets:', error);
    return res.status(400).json({ error: 'Error en la conexión', status: false });
  }
});

router.post('/createIndicator', async (req, res) => {
  try {
    const { nombre, oficinaEscuela, responsable, coequipero, meta2024, meta2025, meta2026, tipoOficinaEscuela, id_obj_dec } = req.body;
    const year = new Date().getFullYear();
    const idEscOfi = oficinaEscuela;

    const sheets = google.sheets({ version: 'v4', auth: jwtClient });

    // Obtener el nombre de la escuela u oficina de la hoja ESC_OFI
    const escOfiResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: '1sp9G8A6-hPUtnmfK7jSpAqQfoMzKR3kmYGYzhOAC6vM', 
      range: 'ESC_OFI!A1:E1000', 
    });

    const escOfiValues = escOfiResponse.data.values;
    const escOfi = escOfiValues.find(row => row[0] === idEscOfi); 

    if (!escOfi) {
      return res.status(404).json({ status: false, message: 'Escuela u oficina no encontrada' });
    }

    const nombreEscuelaOficina = escOfi[1]; 
    const tipo = escOfi[2]; 

    // Crear el nombre correcto del archivo de Google Sheets
    const tipoArchivo = tipo === 'Escuela' ? 'Esc' : 'Ofi';
    const sheetName = `${year}MDE - ${tipoArchivo} ${nombreEscuelaOficina} - IndicadorNo_${Math.floor(Math.random() * 1000)}`;

    const drive = google.drive({ version: 'v3', auth: jwtClient });

    // Crear el archivo en Google Drive
    const fileMetadata = {
      name: sheetName,
      mimeType: 'application/vnd.google-apps.spreadsheet',
      parents: ['1wBWPuy0TH3rNnQvGA9mEDgYLYxq48HgQ'], // ID de la carpeta 
    };

    const file = await drive.files.create({
      resource: fileMetadata,
      fields: 'id',
    });

    const spreadsheetId = file.data.id;
    const urlIndicador = `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;

    // Obtener los IDs actuales de la hoja INDICADORES
    const indicadoresSheetId = '1sp9G8A6-hPUtnmfK7jSpAqQfoMzKR3kmYGYzhOAC6vM'; // ID de la hoja INDICADORES
    const indicadoresResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: indicadoresSheetId,
      range: 'INDICADORES!A1:A',
    });

    const currentIds = indicadoresResponse.data.values.slice(1).map(row => parseInt(row[0]));
    const newId = Math.max(...currentIds) + 1;

    // Agregar el nuevo indicador a la hoja INDICADORES (ID del objetivo decanato en `id_obj_dec`, no en `id_obj2`)
    const newIndicator = [newId, nombre, '', id_obj_dec, responsable, coequipero, urlIndicador, idEscOfi];

    await sheets.spreadsheets.values.append({
      spreadsheetId: indicadoresSheetId,
      range: 'INDICADORES!A1:H1',
      valueInputOption: 'RAW',
      resource: {
        values: [newIndicator],
      },
    });

    // Calcular la meta trienal
    const metaTrienio = parseFloat(meta2024) + parseFloat(meta2025) + parseFloat(meta2026);

    // Obtener los IDs actuales de la hoja METAS
    const metasResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: indicadoresSheetId,
      range: 'METAS!A1:A',
    });

    const currentMetaIds = metasResponse.data.values.slice(1).map(row => parseInt(row[0]));
    const newMetaId = Math.max(...currentMetaIds) + 1;

    // Agregar las metas a la hoja METAS
    const newMeta = [newMetaId, newId, meta2024, '', meta2025, '', meta2026, '', metaTrienio, ''];

    await sheets.spreadsheets.values.append({
      spreadsheetId: indicadoresSheetId,
      range: 'METAS!A1:J1',
      valueInputOption: 'RAW',
      resource: {
        values: [newMeta],
      },
    });

    return res.status(200).json({ status: true, message: 'Indicador y metas creados con éxito', url: urlIndicador });
  } catch (error) {
    console.error('Error creando indicador y metas:', error);
    return res.status(500).json({ status: false, message: 'Error al crear el indicador y las metas', error });
  }
});


app.use(router);  

app.listen(PORT, () => {
  console.log(`Servidor backend escuchando en el puerto ${PORT}`);
});
