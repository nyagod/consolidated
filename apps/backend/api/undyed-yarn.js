import fs from 'fs';
import path from 'path';
import xlsx from 'xlsx';

// Get the absolute path to the Excel file in the 'api' folder
const filePath = path.join(process.cwd(), 'apps', 'backend/api/CONSOLIDATED REPORT OCTOBER 2024.xlsx');

// Function to read the Excel file asynchronously
async function readExcelFile(filePath) {
    try {
        const fileData = await fs.promises.readFile(filePath);
        const workbook = xlsx.read(fileData, { type: 'buffer' });
        return workbook;
    } catch (error) {
        console.error('Error reading file:', error);
        throw error;
    }
}

// Load the workbook and access the 'Consolidated' sheet
async function loadData() {
    const workbook = await readExcelFile(filePath);
    const sheet = workbook.Sheets['Consolidated'];

    // Limit the range by specifying rows and columns (for example, A1:F50)
    const range = 'A1:F50'; // Adjust this range based on your needs
    const rawData = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: null, range: range });

    return rawData;
}

// Find the section start index
function findSection(keyword, rawData) {
    return rawData.findIndex(row => row[0] && row[0].toString().includes(keyword));
}

// Convert section to dictionary format
function arrayToDictionary(data) {
    const headers = data[1];
    return data.slice(2).filter(row => row.length > 0).map(row => {
        return headers.reduce((obj, header, index) => {
            obj[header] = row[index] || null;
            return obj;
        }, {});
    });
}

// API handler
export default async function handler(req, res) {
    try {
        const rawData = await loadData();

        const undyedYarnStart = findSection('UNDYED YARN / STOCK', rawData);
        const availableSectionStart = findSection('AVAILABLE SECTION', rawData);

        if (undyedYarnStart === -1 || availableSectionStart === -1) {
            return res.status(404).json({ message: 'Sections not found in the data.' });
        }

        const undyedYarnData = arrayToDictionary(rawData.slice(undyedYarnStart, availableSectionStart));

        return res.status(200).json(undyedYarnData);
    } catch (error) {
        console.error('Error processing data:', error);
        return res.status(500).json({ message: 'Internal Server Error', error: error.message });
    }
}
