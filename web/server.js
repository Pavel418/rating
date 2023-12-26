const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const e = require('express');

const app = express();
const port = 3000; // Choose your desired port number

app.use(cors());

app.get('/api/students', async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('output.xlsx');
    const previous = new ExcelJS.Workbook();
    await previous.xlsx.readFile('previous.xlsx');

    const worksheet = workbook.getWorksheet(1);
    const prevWorksheet = previous.getWorksheet(1);

    const data = [];

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber !== 1) {
        const [id, studentName, faculty, average] = row.values.slice(1);

        let prevRow;
        for (let rowNumber = 1; rowNumber <= prevWorksheet.rowCount; rowNumber++) {
            let row = prevWorksheet.getRow(rowNumber);
            if (row.getCell(2).value.trim() === studentName.trim()) {
                prevRow = row;
                break;
            }
        }

        let difference = 0;
        if (prevRow) {
          const prevId = prevRow.getCell(1).value;
          difference = prevId - id;
        }

        data.push({
          id,
          studentName,
          faculty,
          average,
          difference
        });
      }
    });

    res.json(data);
  } catch (error) {
    console.error('Error reading Excel files:', error);
    res.status(500).json({ error: 'Unable to read Excel files' });
  }
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
