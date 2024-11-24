const xlsx = require('xlsx');

// Load the workbook and access the 'Consolidated' sheet
const workbook = xlsx.readFile('CONSOLIDATED REPORT OCTOBER 2024.xlsx');
const sheet = workbook.Sheets['Consolidated'];
const rawData = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: null });

// Find section start indices
function findSection(keyword) {
    return rawData.findIndex(row => row[0] && row[0].toString().includes(keyword));
}

const dyedStoreStart = findSection('DYED STORE');
const totalColorUsedStart = findSection('TOTAL COLOR USED');

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

const dyedStoreData = arrayToDictionary(rawData.slice(dyedStoreStart, totalColorUsedStart));

// API handler
export default function handler(req, res) {
    res.status(200).json(dyedStoreData);
}
