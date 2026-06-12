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

// MAIN ROUTE: Receives multiple excel files, reads data, merges them intelligently, and returns a Master File + Merged Data JSON
app.post('/api/merge', upload.array('excelFiles'), (req, res) => {
    try {
        if (!req.files || req.files.length === 0) {
            return res.status(400).json({ success: false, message: 'No files uploaded.' });
        }

        // Object map to group and combine student rows by their unique PRN_NO
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
            
            // PROCESS EACH ROW INTELLIGENTLY
            jsonData.forEach((row) => {
                // Find the PRN key dynamically (handles whitespace or lowercase variations like 'prn_no')
                const prnKey = Object.keys(row).find(k => k.trim().toUpperCase() === 'PRN_NO');
                
                if (prnKey && row[prnKey] !== undefined) {
                    const prnValue = row[prnKey].toString().trim();

                    if (!studentMap[prnValue]) {
                        // If PRN is encountered for the first time, create a new entry
                        studentMap[prnValue] = { ...row };
                    } else {
                        // If PRN already exists, merge new columns into the same row without creating a duplicate line!
                        studentMap[prnValue] = { ...studentMap[prnValue], ...row };
                    }
                }
            });

            // Synchronously delete the file from 'uploads' folder to free local disk storage space
            if (fs.existsSync(filePath)) {
                fs.unlinkSync(filePath);
            }
        });

        // Convert our intelligently grouped student map back into a clean flat array list
        const masterData = Object.values(studentMap);

        // Create a completely new Blank Workbook Object for our final master deliverable
        const newWorkbook = XLSX.utils.book_new();
        
        // Convert our compiled JavaScript Master Array data back into binary Worksheet structures
        const newWorksheet = XLSX.utils.json_to_sheet(masterData);
        
        // Append the sheet data into our new workbook container given a tab title name
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Master Student Report');

        // Setup a unique output name path identifier destination inside memory system
        const outputFileName = `Master_Student_Data_${Date.now()}.xlsx`;
        const uploadDir = path.join(__dirname, 'uploads');
        const outputPath = path.join(uploadDir, outputFileName);

        // Commit file system write execution commands out onto the local memory drive directory
        XLSX.writeFile(newWorkbook, outputPath);

        // Sending back the direct download link AND the raw merged data array for frontend table preview feature
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