const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const app = express();

app.use(express.static('../public'));
app.use(express.json());

// Configure multer
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, 'uploads/'); // Files will be saved in 'uploads' folder
    },
    filename: (req, file, cb) => {
        cb(null, file.originalname); // Unique filename
    }
});

const upload = multer({ storage: storage });

// Create uploads directory if it doesn’t exist
if (!fs.existsSync('uploads')) {
    fs.mkdirSync('uploads');
}

let latestFile = null;

const middleware = (req, res, next) => {
    console.log("1234567890")
    next();
}

const uploadRoute = async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({
                msg: 'No file uploaded'
            });
        }
        latestFile = req.file.path;
        res.status(200).json({
            msg: 'success',
            filename: req.file.filename,
            path: req.file.path
        });
    } catch (error) {
        res.status(500).json({
            msg: 'failed',
            error: error.message
        });
    }    
}

app.post('/upload', upload.single('excelFile'), uploadRoute)

app.get('/get-new-file', (req, res) => {
    try {
        if (!latestFile || !fs.existsSync(latestFile)) {
            return res.status(404).json({
                msg: 'No file available'
            });
        }
    
        const workbook = XLSX.readFile(latestFile);
        const sheetName = workbook.SheetNames[0]; // Get the first sheet
        const worksheet = workbook.Sheets[sheetName];
        const excelData = XLSX.utils.sheet_to_json(worksheet);
    
        const record = new Map();
    
        excelData.map((row) => {
            if (record.has(row.Segment)) {
                record.set(row.Segment, [...record.get(row.Segment), row])
            } else {
                record.set(row.Segment, [row])
            }        
        })
    
        let index = 0;
        let newFileName = null;
        let newFilePath = null;

        const newWorkbook = XLSX.utils.book_new();

        for (const [key, value] of record) {
            // Create a new workbook and worksheet
            const worksheet = XLSX.utils.json_to_sheet(value);
    
            // Add the worksheet to the workbook
            XLSX.utils.book_append_sheet(newWorkbook, worksheet, `Sheet${index+1}`);
            index++;
        }

        if (!fs.existsSync('generated')) {
            fs.mkdirSync('generated');
        }
        
        // Define the output file path
        const fileName = `new-file.xlsx`;
        const filePath = `generated/${fileName}`;

        // Write the file to disk
        XLSX.writeFile(newWorkbook, filePath);

        newFileName = fileName
        newFilePath = filePath


        console.log(newFilePath, filePath)

        if (!newFilePath || !fs.existsSync(newFilePath)) {
            return res.status(404).json({ msg: 'No generated file available' });
        }

        res.download(newFilePath, path.basename(newFilePath), (err) => {
            if (err) {
                res.status(500).json({ msg: 'Error sending file', error: err.message });
            }
        });
    } catch (error) {
        res.status(500).json({
            msg: 'failed',
            error: error.message
        });
    }   
})

app.listen(3000, () => {
    console.log('Server started!!!')
})
