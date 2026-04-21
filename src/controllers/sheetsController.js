import {sheetRanges} from '../config/sheetRanges.js';
import { config } from 'dotenv';
import { google } from 'googleapis';
import { jwtClient } from '../config/google.js';
import { sheetValuesToObject } from '../utils/utils.js';
config();



export const getAllSheetsData = async (req, res) => {
  try {
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });
    const spreadsheetId = process.env.spreadsheet;

    // Promesas para cada hoja
    const dataPromises = Object.entries(sheetRanges).map(async ([sheetName, range]) => {
      const fullRange = `${sheetName}!${range}`;
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: fullRange,
      });
      return { [sheetName]: response.data.values };
    });

    // Espera a que todas las promesas se resuelvan
    const allDataArray = await Promise.all(dataPromises);

    // Unifica los resultados en un solo objeto
    const allData = Object.assign({}, ...allDataArray);
    const allDataWithObjects = {};

    for (const [sheetName, values] of Object.entries(allData)) {
      allDataWithObjects[sheetName] = sheetValuesToObject(values);
    }
    res.status(200).json({ status: true, data: allDataWithObjects }); //
  } catch (error) {
    console.error('Error obteniendo datos de todas las hojas:', error);
    res.status(400).json({ status: false, error });
  }
};