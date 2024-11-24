const xlsx = require('xlsx');

// Load the workbook and access the 'Consolidated' sheet
const workbook = xlsx.readFile('./CONSOLIDATED REPORT OCTOBER 2024.xlsx');
const sheet = workbook.Sheets['Consolidated'];
const rawData = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: null });

// Find the section start index
function findSection(keyword) {
    return rawData.findIndex(row => row[0] && row[0].toString().includes(keyword));
}

const undyedYarnStart = findSection('UNDYED YARN / STOCK');
const availableSectionStart = findSection('AVAILABLE SECTION');

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

const undyedYarnData = arrayToDictionary(rawData.slice(undyedYarnStart, availableSectionStart));

// API handler
export default function handler(req, res) {
    res.status(200).json(undyedYarnData);
}
