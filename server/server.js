const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const app = express();

app.use(express.static(path.join(__dirname, "public")));
app.use(express.json());

// Configure multer to use memory storage instead of disk storage
const upload = multer({ storage: multer.memoryStorage() });

app.post('/upload', upload.single('excelFile'), async (req, res) => {
    try {
        // Check if file exists
        if (!req.file) {
            return res.status(400).json({
                msg: 'No file uploaded'
            });
        }
        
        const sortedBy = req.body.sortedBy || 'Segment';

        // Read the Excel file from buffer (memory)
        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const excelData = XLSX.utils.sheet_to_json(worksheet);

        if (excelData.length > 0 && !excelData[0].hasOwnProperty(sortedBy)) {
            return res.status(400).json({
                msg: `Invalid segment property: '${sortedBy}' not found in Excel data`
            });
        }

        // Process data into segments
        const record = new Map();
        excelData.forEach((row) => {
            if (record.has(row[sortedBy])) {
                record.set(row[sortedBy], [...record.get(row[sortedBy]), row]);
            } else {
                record.set(row[sortedBy], [row]);
            }
        });

        // Create new workbook in memory
        const newWorkbook = XLSX.utils.book_new();
        let index = 0;

        // Add sheets to new workbook
        for (const [key, value] of record) {
            const worksheet = XLSX.utils.json_to_sheet(value);
            XLSX.utils.book_append_sheet(newWorkbook, worksheet, key);
            index++;
        }

        // Generate Excel file in memory as buffer
        const excelBuffer = XLSX.write(newWorkbook, {
            bookType: 'xlsx',
            type: 'buffer',
            compression: true
        });

        // Set response headers for file download
        res.setHeader('Content-Disposition', 'attachment; filename="processed-file.xlsx"');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        
        // Send the buffer directly as response
        res.send(excelBuffer);

    } catch (error) {
        res.status(500).json({
            msg: 'failed',
            error: error.message
        });
    }
});

app.listen(3000, () => {
    console.log('Server started!!!');
});