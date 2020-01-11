const pdf2table = require('pdf2table');
const fs = require('fs');
const XlsxPopulate = require('xlsx-populate');


function parsePDF() {
    fs.readFile('./para01.pdf', function(err, buffer) {
        if (err) return console.log(err);

        pdf2table.parse(buffer, function (err, rows, rowsdebug) {
            if (err) return console.log(err);

            console.log(rows);
        })
    });
}

function exportToExcel(data) {
    XlsxPopulate.fromFileAsync('./benefix-template.xlsx')
        .then(workbook => {
            const row1 = workbook.sheet('Blank Upload Template').range("A2:AZ2");
            row1.value([data])
            
            return workbook.toFileAsync("./benefix-template.xlsx");
        })
}

parsePDF();
