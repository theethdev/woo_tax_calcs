const moment = require('moment');
const XLSX = require('xlsx');

function readExcel(filePath) {
    return new Promise((resolve) => {
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const rawTrades = XLSX.utils.sheet_to_json(worksheet);

        const trades = rawTrades.map((trade) => {
            const excelTimestampFilled = trade['Filled Time'];
            const filledTime = moment((excelTimestampFilled - (25567 + 2)) * 86400 * 1000).format('YYYY-MM-DD HH:mm:ss');

            const excelTimestampCreated = trade['Create Time'];
            const createTime = moment((excelTimestampCreated - (25567 + 2)) * 86400 * 1000).format('YYYY-MM-DD HH:mm:ss');

            return { ...trade, 'Filled Time': filledTime, 'Create Time': createTime };
        });

        resolve(trades);
    });
}

function calculatePnL(trades) {
    const positions = {};
    const pnlByDateTime = {};
    let totalPnL = 0;

    trades.sort((a, b) => new Date(a['Filled Time']) - new Date(b['Filled Time']));

    trades.forEach((trade) => {
        const {
            'Order ID': orderId,
            'Filled Time': filledTime,
            'Instrument': instrument,
            'Side': side,
            'Price': price,
            'Quantity': quantity,
            'Executed': executed,
            'Average Price': avgPrice,
            'Amount': executedAmount,
            'Status': status,
            'Total Fee': totalFee,
            'Fee Token': feeToken,
        } = trade;

        if (side !== 'BUY' && side !== 'SELL') {
            console.warn(`Invalid Side value "${side}" for Order ID ${trade['Order ID']}. Skipping this trade.`);
            return;
        }

        if (!positions[instrument]) {
            positions[instrument] = [];
        }

        // var date = moment(filledTime).format('YYYY-MM-DD');

        const positionKey = side === 'BUY' ? 'long' : 'short';
        const oppositeKey = side === 'BUY' ? 'short' : 'long';

        if (!positions[instrument]) {
            positions[instrument] = {};
        }
        if (!positions[instrument][positionKey]) {
            positions[instrument][positionKey] = [];
        }
        if (!positions[instrument][oppositeKey]) {
            positions[instrument][oppositeKey] = [];
        }

        if (side === 'BUY' || side === 'SELL') {
            let remainingQuantity = +executed;
            let tradePnL = 0;

            while (remainingQuantity > 0 && positions[instrument][oppositeKey].length > 0) {
                const position = positions[instrument][oppositeKey].shift();
                const closedQuantity = Math.min(remainingQuantity, position.quantity);
                const pnl = (positionKey === 'long' ? position.avgPrice - avgPrice : avgPrice - position.avgPrice) * closedQuantity - totalFee;
                tradePnL += pnl;
                remainingQuantity -= closedQuantity;

                if (position.quantity > closedQuantity) {
                    positions[instrument][oppositeKey].unshift({
                        quantity: position.quantity - closedQuantity,
                        avgPrice: position.avgPrice,
                        fee: position.fee,
                    });
                }
            }

            if (remainingQuantity > 0) {
                positions[instrument][positionKey].push({
                    quantity: remainingQuantity,
                    avgPrice: +avgPrice,
                    fee: +totalFee,
                });
            }

            totalPnL += tradePnL;

            const dateTime = moment(filledTime).format('YYYY-MM-DD HH:mm:ss');
            if (!pnlByDateTime[dateTime]) {
                pnlByDateTime[dateTime] = 0;
            }
            pnlByDateTime[dateTime] += tradePnL;
        }
    });


    // Filter out datetime entries with PnL value of 0
    const filteredPnlByDateTime = Object.fromEntries(
        Object.entries(pnlByDateTime).filter(([_, pnl]) => pnl !== 0)
    );

    return { pnlByDateTime: filteredPnlByDateTime, totalPnL };
}

function writeExcel(data, filename) {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, 'PnL');
    XLSX.writeFile(wb, filename);
}


(async () => {
    console.clear();
    // XXXXX is the woo subaccount id
    preparePnLFile('XXXXX');
})();


async function preparePnLFile(sAccId) {
    const filePath = `./Input/Filled-Order-History-${sAccId}.xlsx`;
    var trades = await readExcel(filePath);

    trades = trades.filter((trade) => trade.Instrument.slice(0, 4) === 'PERP');
    const { pnlByDateTime, totalPnL } = calculatePnL(trades);

    // console.log('PnL by Date and Time:');
    // for (const [dateTime, pnl] of Object.entries(pnlByDateTime)) {
    //     console.log(`  ${dateTime}: ${pnl.toFixed(2)}`);
    // }

    console.log(`\nTotal PnL: ${totalPnL.toFixed(2)}`);

    const pnlData = Object.entries(pnlByDateTime).map(([date, pnl]) => ({
        'Koinly Date': date,
        Amount: pnl,
        Currency: 'USDT',
        Label: 'realized gain',
    }));

    writeExcel(pnlData, `./Output/Perp_PnL_${sAccId}.xlsx`);
}