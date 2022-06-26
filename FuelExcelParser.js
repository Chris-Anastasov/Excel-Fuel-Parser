import ExcelJS from 'exceljs';
import readline from 'readline';
import fs from 'fs';

const wb = new ExcelJS.Workbook();
const fileName = 'TransactionsExcelFixed.xlsx';
let colNames = []; // Съдържа имената на колоните
let ws;
// Съдържат номерата на колоните
let lpnCol;
let dateCol;
let hourCol;
let quantityCol;

// Функция за събиране на данни чрез user-a
function askQuestion(query) {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    })

    return new Promise(resolve => rl.question(query, ans => {
        rl.close();
        resolve(ans);
    }))
}

// Инициализиране на библиотеката за работа с ексел
wb.xlsx.readFile(fileName).then(async () => {
        ws = wb.getWorksheet('Worksheet');
        const firstRow = ws.getRow(1);

        firstRow.eachCell(c => {
            colNames.push(`\n${c.col}.${c.value}`)
        })

    }).then(function () {
        getColumns();
    })
    .catch(err => {
        console.log(err.message);
    });

// Взима всички нужни параметри, поради разлики в екселските документи
const getColumns = async () => {
    const dateFormats = ['1. dd mm yyyy', '\n2. mm dd yyyy', '\n3. yyyy mm dd', '\n4. yyyy dd mm\n'];

    await askQuestion(`${colNames}\nНапишете номера на колоната съдържаща рег. номер: `).then(lpnVal => {
        lpnCol = ws.getColumn(parseInt(lpnVal));
        askQuestion("Номер на колоната с ДАТА: ").then(dateVal => {
            dateCol = ws.getColumn(parseInt(dateVal));
            askQuestion("Номер на колоната с ЧАС: ").then(hourVal => {
                hourCol = ws.getColumn(parseInt(hourVal));
                askQuestion("Номер на колоната с КОЛИЧЕСТВО: ").then(qtyVal => {
                    quantityCol = ws.getColumn(parseInt(qtyVal));
                    askQuestion(`Формат на датата: \n${dateFormats}`).then(formatSyle => {
                        collectVehicleData(formatSyle);
                    })
                })
            })
        });
    })
}


// Събира данните и ги сортира за експорт
const collectVehicleData = (formatSyle) => {
    let uniqueVehicles = {}
    let vehiclesData = {
        lpn: [],
        dateTime: [],
        qty: []
    };

    // Събира данните от колоните и пълни обект
    lpnCol.eachCell(function (cell, rowNumber) {
        let formatedDateTime;
        if (cell.value !== null && rowNumber !== 1) {
            const row = ws.getRow(rowNumber);
            const dateVal = row.getCell(dateCol._number).value;
            const hourVal = row.getCell(hourCol._number).value;
            const qtyVal = row.getCell(quantityCol._number).value;

            //Спрямо различните формати, разбива датата и я подрежда правилно
            if (formatSyle == 1) {
                const day = dateVal.substring(0, 2);
                const month = dateVal.substring(3, 5);
                const year = dateVal.substring(6, 10);
                formatedDateTime = `${year}-${month}-${day} ${hourVal}`;
            } else if (formatSyle == 2) {
                const month = dateVal.substring(0, 2);
                const day = dateVal.substring(3, 5);
                const year = dateVal.substring(6, 10);
                formatedDateTime = `${year}-${month}-${day} ${hourVal}`;
            } else if (formatSyle == 3) {
                const year = dateVal.substring(0, 4);
                const month = dateVal.substring(5, 7);
                const day = dateVal.substring(8, 10);
                formatedDateTime = `${year}-${month}-${day} ${hourVal}`;
            } else if (formatSyle == 4) {
                const year = dateVal.substring(0, 4);
                const day = dateVal.substring(5, 7);
                const month = dateVal.substring(8, 10);
                formatedDateTime = `${year}-${month}-${day} ${hourVal}`;
            }

            vehiclesData.lpn.push({
                rId: rowNumber,
                value: cell.value
            });
            vehiclesData.dateTime.push({
                rId: rowNumber,
                value: formatedDateTime
            });
            vehiclesData.qty.push({
                rId: rowNumber,
                value: qtyVal
            });
        }
    });

    vehiclesData.lpn.forEach((el, i) => {
        const dateTime = vehiclesData.dateTime.find(e => e.rId == el.rId).value;
        const qty = vehiclesData.qty.find(e => e.rId == el.rId).value;
        // При вече създаден обект, от долу, допълва данните
        if (uniqueVehicles.hasOwnProperty(el.value)) {
            uniqueVehicles[el.value].lpn.push(el.value)
            uniqueVehicles[el.value].dateTime.push(dateTime)
            uniqueVehicles[el.value].qty.push(qty)
        } else {
            // Създава нов обект с данни
            uniqueVehicles[el.value] = {
                lpn: [el.value],
                dateTime: [dateTime],
                qty: [qty]
            }
        }
    })

    exportToExcel(uniqueVehicles);
}

function exportToExcel(uniqueVehicles) {
    for (const vehicleLpn in uniqueVehicles) {
        if (Object.hasOwnProperty.call(uniqueVehicles, vehicleLpn)) {
            const vehicleData = uniqueVehicles[vehicleLpn];

            const privateWb = new ExcelJS.Workbook();
            const privateWs = privateWb.addWorksheet('Worksheet');

            const headers = [{
                    header: 'Регистрационен номер',
                    key: 'lpn',
                    width: 20
                },
                {
                    header: 'Дата и час на зарежедане',
                    key: 'dateTime',
                    width: 20
                },
                {
                    header: 'Количество заредено гориво',
                    key: 'qtyVolume',
                    width: 25
                },
                {
                    header: 'Наличност в края на зареждането',
                    key: 'tankVolume',
                    width: 25
                },
            ]

            privateWs.columns = headers;

            vehicleData.lpn.forEach((el, i) => {
                privateWs.addRow([el, vehicleData.dateTime[i], vehicleData.qty[i], vehicleData.qty[i]])
            })

            fs.existsSync("Fuels folder") || fs.mkdirSync("Fuels folder");

            privateWb.xlsx.writeFile("/d" + vehicleLpn + '.xlsx').then(() => {
                console.log('file created');
            });
        }
    }
}