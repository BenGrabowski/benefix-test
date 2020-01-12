const pdf2table = require('pdf2table');
const fs = require('fs');
const XlsxPopulate = require('xlsx-populate');

function parsePDF() {
    fs.readFile('./para01.pdf', function(err, buffer) {
        if (err) return console.log(err);

        pdf2table.parse(buffer, function (err, rows, rowsdebug) {
            if (err) return console.log(err);

            getData(rows);

            // let plan = rows.splice(0, 24);

            // createRow(plan, rows);
            
            // const startDate = rows[0].toString().slice(28,37);

            // const endDate = rows[0].toString().slice(41, 50);

            // const planName = rows[2][3];

            // const ratingArea = rows[1][1].slice(4);

            // const state = rows[1][1].slice(0,2);

            // let rateArray = [];
            
            // const rateSection = rows.slice(4, 19);

            // rateSection.forEach(row => {
            //     row.map(item => {
            //         if (item.length > 4) {
            //             rateArray.push(Number(item));
            //         }
            //     })
            // })

            // const rates = rateArray.sort((a,b) => a - b);
            // const row = [startDate, endDate, planName, state, ratingArea, rates[0], ...rates, rates[44]];
            
            // exportToExcel(row);
        })
    });
}

function getData(data) {
    if (!data) {
        return;
    }
    
    let plan = data.splice(0, 24);
    createRow(plan, data);
}

let excelRows = [];

function createRow(plan, data) {
    if (data.length === 0) {
        // console.log(excelRows);
        // return;
        exportToExcel(excelRows);
    }
    // console.log(plan[0]);
    const startDate = plan[0].toString().slice(28,37);
    const endDate = plan[0].toString().slice(41, 50);
    const planName = plan[2][3];
    const ratingArea = plan[1][1].slice(4);
    const state = plan[1][1].slice(0,2);

    let rateArray = [];
    
    const rateSection = plan.slice(4, 19);

    rateSection.forEach(row => {
        row.map(item => {
            if (item.length > 4) {
                rateArray.push(Number(item));
            }
        })
    })

    const rates = rateArray.sort((a,b) => a - b);
    const row = [startDate, endDate, planName, state, ratingArea, rates[0], ...rates, rates[44]];
    
    excelRows.push(row);
    getData(data);
}

function exportToExcel(data) {
    XlsxPopulate.fromFileAsync('./benefix-template.xlsx')
        .then(workbook => {
            // const row1 = workbook.sheet('Blank Upload Template').cell('A2');
            // row1.value(data)

            workbook.sheet('Blank Upload Template').cell('A2').value(data);
            
            return workbook.toFileAsync("./benefix-template.xlsx");
        })
}

parsePDF();