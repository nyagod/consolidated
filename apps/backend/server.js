const express = require('express');
const xlsx = require('xlsx');
const cors = require('cors');

const app = express();
module.exports = app; // CommonJS syntax
 // Change to ES module syntax
const PORT = 3000;
app.use(cors({
    origin: '*', // Allow requests only from your frontend domain (localhost:3001)
    methods: ['GET', 'POST', 'PUT', 'DELETE'],
    credentials: true // Optional, if you need to send cookies with the requests
}));

// Load the workbook and access the 'Consolidated' sheet
const workbook = xlsx.readFile('CONSOLIDATED REPORT OCTOBER 2024.xlsx'); // Replace with the actual file name
const sheet = workbook.Sheets['Consolidated'];

// Convert the entire sheet to JSON with arrays of rows
const rawData = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: null });

// Function to find a section start based on keyword
function findSection(keyword) {
    return rawData.findIndex(row => row[0] && row[0].toString().includes(keyword));
}

// Helper function to convert section arrays to dictionary format
function arrayToDictionary(data) {
    const headers = data[1]; // Assuming the second row contains headers (date columns + other labels)
    return data.slice(2).filter(row => row.length > 0).map(row => {
        // Merge headers with the data
        return headers.reduce((obj, header, index) => {
            obj[header] = row[index] || null; // Add row data as key-value pairs
            return obj;
        }, {});
    });
}

// Separate each section and convert to dictionary format
const undyedYarnStart = findSection('UNDYED YARN / STOCK');
const availableSectionStart = findSection('AVAILABLE SECTION');
const dyedStoreStart = findSection('DYED STORE');
const totalColorUsedStart = findSection('TOTAL COLOR USED');
const totalOrderReceivedStart = findSection('TOTAL ORDER RECEIVED');
const partyTotalStart = findSection('PARTY');

// Separate and convert sections to dictionaries
const sections = {
    undyedYarnData: arrayToDictionary(rawData.slice(undyedYarnStart, availableSectionStart)),
    availableSectionData: arrayToDictionary(rawData.slice(availableSectionStart, dyedStoreStart)),
    dyedStoreData: arrayToDictionary(rawData.slice(dyedStoreStart, totalColorUsedStart)),
    totalColorUsedData: arrayToDictionary(rawData.slice(totalColorUsedStart, totalOrderReceivedStart)),
    totalOrderReceivedData: arrayToDictionary(rawData.slice(totalOrderReceivedStart, partyTotalStart)),
    partyTotalData: arrayToDictionary(rawData.slice(partyTotalStart))
};

// Define routes to serve each section
app.get('/',(req,res) => res.json("hello message"));
app.get('/undyed-yarn', (req, res) => res.json(sections.undyedYarnData));
app.get('/available-section', (req, res) => res.json(sections.availableSectionData));
app.get('/dyed-store', (req, res) => res.json(sections.dyedStoreData));
app.get('/total-color-used', (req, res) => res.json(sections.totalColorUsedData));
app.get('/total-order-received', (req, res) => res.json(sections.totalOrderReceivedData));
app.get('/party-total', (req, res) => res.json(sections.partyTotalData));

// Start the Express server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
