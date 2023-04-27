const XLSX = require('xlsx');
const moment = require('moment-timezone');

process.env.TZ = 'UTC';
console.clear();

// Read Excel file
function readExcelFile(filePath) {
    const sInterestSheet = 'Interest History';
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[sInterestSheet];
    const rawData = XLSX.utils.sheet_to_json(sheet);

    const data = rawData.map((line) => {
        const excelTimestampFilled = line['Time'];
        const filledTime = moment((excelTimestampFilled - (25567 + 2)) * 86400 * 1000).format('YYYY-MM-DD HH:mm:ss');
        return { ...line, 'Time': filledTime };
    });

    const sortedData = data.sort((a, b) => new Date(a.Time) - new Date(b.Time));
    return sortedData.filter(row => row.Action === 'LOAN');
}


// Calculate loan taken/repaid
function calculateLoan(rows) {
    let loanHistory = [];
    let previousLoan = 0;
    let totalLoan = 0;

    if (rows.length === 0)
        return loanHistory;

    rows.forEach(row => {
        let currentLoan = row['Borrow Quantity'];
        let diff = currentLoan - previousLoan;

        if (diff !== 0) {
            loanHistory.push({
                datetime: row.Time,
                amount: diff
            });
            totalLoan += diff;
            previousLoan = currentLoan;
        }
    });

    // Add the repayment of the balance amount
    const lastEntryDatetime = moment(rows[rows.length - 1].Time);
    const repaymentDatetime = lastEntryDatetime.add(1, 'hour').format('YYYY-MM-DD HH:mm:ss');
    loanHistory.push({
        datetime: repaymentDatetime,
        amount: -totalLoan
    });

    return loanHistory;
}

function writeToExcel(loanHistory, outputFilePath) {
    const data = loanHistory.map(entry => {
        const loanDatetime = moment(entry.datetime);
        const isLoan = entry.amount > 0;

        if (isLoan) {
            loanDatetime.subtract(1, 'hour');
        } else {
            loanDatetime.add(1, 'minute');
        }

        return {
            Currency: 'USDT',
            'Koinly Date': loanDatetime.format('YYYY-MM-DD HH:mm:ss'),
            Amount: entry.amount,
            Description: isLoan ? 'LOAN' : 'REPAYMENT'
        };
    });

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'LoanHistory');

    XLSX.writeFile(workbook, outputFilePath);
}



// Main function to run the program
function start(sAccId) {
    const inputFilePath = `./Input/Wallet-History-${sAccId}.xlsx`;
    const outFilePath = `./Output/Loans-${sAccId}.xlsx`;
    const data = readExcelFile(inputFilePath);

    const loanHistory = calculateLoan(data);

    console.log('Datetime, Amount of Loan Taken/Repayed');
    loanHistory.forEach(entry => {
        // console.log(`${entry.datetime}, ${entry.amount}`);
    });

    if (loanHistory.length > 0)
        writeToExcel(loanHistory, outFilePath);
    else
        console.log(`No loan history found for ${sAccId}`);
}

start('XXXXXX');
