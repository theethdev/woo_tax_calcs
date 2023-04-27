const XLSX = require('xlsx');
const moment = require('moment-timezone');

process.env.TZ = 'UTC';

(async () => {
    start('XXXXXX'); // XXXXXX is the account ID
})();

async function start(sAccId) {
    const sFundingSheet = 'Funding Fee History';
    const sInterestSheet = 'Interest History';
    const inFilePath = `./Input/Wallet-History-${sAccId}.xlsx`;
    var aggregatedInterestHistory = [];
    var aggregatedFundingFeeHistory = [];

    const interestHistoryData = await readExcel(inFilePath, sInterestSheet);
    if (interestHistoryData) {
        aggregatedInterestHistory = aggregateInterestHistory(interestHistoryData);
    } else {
        console.error('Interest History sheet not found');
    }

    const fundingFeeHistoryData = await readExcel(inFilePath, sFundingSheet);
    if (fundingFeeHistoryData) {
        aggregatedFundingFeeHistory = aggregateFundingFeeHistory(fundingFeeHistoryData);
    } else {
        console.error('Funding Fee History sheet not found');
    }

    if (interestHistoryData && fundingFeeHistoryData) {
        const mergedData = [...aggregatedInterestHistory, ...aggregatedFundingFeeHistory];
        writeExcel(mergedData, `./Output/Interest_Funding-${sAccId}.xlsx`, sFundingSheet);
    }
}

function aggregateInterestHistory(data) {
    const aggregatedData = [];

    // sort data by time ascending
    data.sort((a, b) => new Date(a['Time']) - new Date(b['Time']));

    data.forEach(row => {
        if (row['Action'] === 'LOAN') {
            const timestamp = new Date(row['Time']);
            const date = new Date(timestamp.getFullYear(), timestamp.getMonth(), timestamp.getDate()).toISOString().split('T')[0];

            const existingEntry = aggregatedData.find(entry => entry.Date === date);

            // Remove the " USDT" part from the Quantity value
            const quantity = parseFloat(row['Quantity'].replace(' USDT', ''));

            if (existingEntry) {
                existingEntry.Amount += quantity;
            } else {
                aggregatedData.push({ Date: date, Amount: quantity });
            }
        }
    });

    const formattedData = aggregatedData.map(entry => ({
        'Koinly Date': entry.Date,
        Amount: entry.Amount,
        Currency: 'USDT',
        Label: 'loan interest',
    }));

    return formattedData;
}

function aggregateFundingFeeHistory(data) {
    const aggregatedData = [];
    data.sort((a, b) => new Date(a['Time']) - new Date(b['Time']));

    data.forEach(row => {
        const timestamp = new Date(row['Time']);
        const date = new Date(timestamp.getFullYear(), timestamp.getMonth(), timestamp.getDate()).toISOString().split('T')[0];

        const existingEntry = aggregatedData.find(entry => entry.Date === date);

        // Extract the Funding Fee Amount
        const fundingFeeAmount = parseFloat(row['Funding Fee Amount']);

        if (existingEntry) {
            existingEntry.Amount += fundingFeeAmount;
        } else {
            aggregatedData.push({ Date: date, Amount: fundingFeeAmount });
        }
    });

    const formattedData = aggregatedData.map(entry => ({
        'Koinly Date': entry.Date,
        Amount: entry.Amount,
        Currency: 'USDT',
        Label: 'realized gain',
    }));

    return formattedData;
}


async function readExcel(filePath, sheetName) {
    return new Promise((resolve) => {
        const workbook = XLSX.readFile(filePath);
        const worksheet = workbook.Sheets[sheetName];
        const rawData = XLSX.utils.sheet_to_json(worksheet);

        const data = rawData.map((line) => {
            const excelTimestampFilled = line['Time'];
            const filledTime = moment((excelTimestampFilled - (25567 + 2)) * 86400 * 1000).format('YYYY-MM-DD HH:mm:ss');
            return { ...line, 'Time': filledTime };
        });

        resolve(data);
    });
}

function writeExcel(data, fileName, sheetName) {
    if (data.length === 0) {
        console.error('No data to write');
        return;
    }
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    XLSX.writeFile(workbook, fileName);
}