const bodyParser = require('body-parser');
const excel = require('exceljs');
const unstream = require('unstream');
const express = require('express');
const app = express();
const http = require('http').createServer(app);

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: false }));

const port = process.env.PORT || 3000;
app.set('port', port);

app.get("/report", (req, res) => {

    res.setHeader('Content-type', 'application/vnd.ms-excell');
    res.setHeader('Content-Transfer-Encoding', 'binary');
    res.setHeader('Content-disposition', 'attachment; filename="report.xlsx"');

    const fileName = 'templates/DAYTIME_TOURISTS.xlsx';

    let workbook = new excel.Workbook();

    workbook.xlsx.readFile(fileName)
        .then(() => {
            let worksheet = workbook.getWorksheet("SAME DAY");

            workbook.xlsx.write(unstream({}, function(buf) {
                res.status(200).send(buf);
            })).catch(err => {
                console.error(err);
            });

        }).catch(err => {
            console.log(err);
        });
});

http.listen(port, () => {
    console.log(`Server running or port ${port}`);
});
