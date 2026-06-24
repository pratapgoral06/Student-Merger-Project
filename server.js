// Import required external npm packages
const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
const PORT = 3000;

// Enable CORS and allow JSON data processing
app.use(cors());
app.use(express.json());

// Serve all static frontend HTML files automatically from the 'public' directory
app.use(express.static('public'));

// Configure Multer storage to save uploaded files inside the 'uploads' directory temporarily
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        const uploadDir = path.join(__dirname, 'uploads/');
        if (!fs.existsSync(uploadDir)) {
            fs.mkdirSync(uploadDir, { recursive: true }); // Create folder dynamically if missing
        }
        cb(null, uploadDir);
    },
    filename: (req, file, cb) => {
        cb(null, Date.now() + '-' + file.originalname); // Unique filename timestamp
    }
});
const upload = multer({ storage: storage });

// MAIN ROUTE: Receives multiple excel files + dynamic primary key, merges them intelligently
app.post('/api/merge', upload.array('excelFiles'), (req, res) => {
    try {
        if (!req.files || req.files.length === 0) {
            return res.status(400).json({ success: false, message: 'No files uploaded.' });
        }

        // Fetch the dynamic primary key parameter from frontend (defaults to 'PRN_NO' if empty)
        const userPrimaryKey = req.body.primaryKey ? req.body.primaryKey.trim().toUpperCase() : 'PRN_NO';

        // Object map to group and combine student rows dynamically by user-defined primary key
        let studentMap = {}; 

        // Loop through each uploaded Excel file sequentially
        req.files.forEach((file) => {
            const filePath = file.path;
            
            // Read the Excel Workbook structure safely
            const workbook = XLSX.readFile(filePath);
            
            // Target the very first sheet inside the workbook file
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert worksheet binary rows into easy JavaScript JSON Objects array
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            // PROCESS EACH ROW INTELLIGENTLY USING THE DYNAMIC KEY
            jsonData.forEach((row) => {
                // Find the target key dynamically (handles whitespace or lowercase variations matching the user input)
                const matchedKey = Object.keys(row).find(k => k.trim().toUpperCase() === userPrimaryKey);
                
                if (matchedKey && row[matchedKey] !== undefined) {
                    const keyValue = row[matchedKey].toString().trim();

                    if (!studentMap[keyValue]) {
                        // If the key value is encountered for the first time, create a new entry
                        studentMap[keyValue] = { ...row };
                    } else {
                        // If the key already exists, merge new columns into the same row without creating duplicates
                        studentMap[keyValue] = { ...studentMap[keyValue], ...row };
                    }
                }
            });

            // Synchronously delete the file from 'uploads' folder to free local disk storage space
            if (fs.existsSync(filePath)) {
                fs.unlinkSync(filePath);
            }
        });

        // Convert our intelligently grouped student map back into a clean flat array list
        const rawMasterData = Object.values(studentMap);

        // --- GLOBAL BASELINE COLUMN SORTING LOGIC ---
        // 1. Extract all unique column headers across all merged rows using a Set
        let allHeaders = new Set();
        rawMasterData.forEach(row => {
            Object.keys(row).forEach(key => allHeaders.add(key));
        });

        // 2. Convert the Set to an Array and sort it alphabetically (A to Z) as a safe baseline sequence
        let sortedHeaders = Array.from(allHeaders).sort();

        // 3. Reconstruct each row object with sorted keys so the frontend preview receives a structured data block
        const masterData = rawMasterData.map(row => {
            let sortedRow = {};
            sortedHeaders.forEach(key => {
                if (row[key] !== undefined) {
                    sortedRow[key] = row[key];
                }
            });
            return sortedRow;
        });

        // Create a completely new Blank Workbook Object for our final master deliverable
        const newWorkbook = XLSX.utils.book_new();
        
        // 4. Convert master array data into binary Worksheet structures matching our baseline sorted schema headers
        const newWorksheet = XLSX.utils.json_to_sheet(masterData, { header: sortedHeaders });
        
        // Append the sheet data into our new workbook container given a tab title name
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Master Student Report');

        // Setup a unique output name path identifier destination inside memory system
        const outputFileName = `Master_Student_Data_${Date.now()}.xlsx`;
        const uploadDir = path.join(__dirname, 'uploads');
        const outputPath = path.join(uploadDir, outputFileName);

        // Commit file system write execution commands out onto the local memory drive directory
        XLSX.writeFile(newWorkbook, outputPath);

        // Sending back the direct download link AND the raw merged data array for frontend customized interactive reordering
        res.json({
            success: true,
            message: 'Files merged successfully!',
            downloadUrl: `/api/download/${outputFileName}`,
            mergedData: masterData
        });

    } catch (error) {
        console.error('Server error during processing:', error);
        res.status(500).json({ success: false, message: 'Internal Server Error during excel merging.' });
    }
});

// DOWNLOAD ROUTE: Allows users to securely download the generated master report
app.get('/api/download/:filename', (req, res) => {
    const fileName = req.params.filename;
    const filePath = path.join(__dirname, 'uploads', fileName);

    if (fs.existsSync(filePath)) {
        res.download(filePath, (err) => {
            if (!err) {
                // Delete the generated master file after successful download transmission to save disk space
                try {
                    fs.unlinkSync(filePath);
                } catch (e) {
                    console.error('Error deleting file:', e);
                }
            }
        });
    } else {
        res.status(404).send('Requested master file expired or not found.');
    }
});

// Start the Express active listener server stack
app.listen(PORT, () => {
    console.log(`Server is running live on: http://localhost:${PORT}`);
});