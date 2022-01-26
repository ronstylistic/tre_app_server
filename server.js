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

    let workbook = new excel.Workbook();

    let worksheet = workbook.addWorksheet("Sheet 1");

    worksheet.columns = [
        { header: "Name of Attraction / Destination", key: "name", width: 20 },
        { header: "Municipality", key: "municipality", width: 20 },
        { header: "Attraction Code", key: "code", width: 20 },
        { header: "Male", key: "province_male", width: 10 },
        { header: "Female", key: "province_female", width: 10 },
        { header: "Total", key: "province_total", width: 10 },
        { header: "Male", key: "other_male", width: 10 },
        { header: "Female", key: "other_female", width: 10 },
        { header: "Total", key: "other_total", width: 10 },
        { header: "Male", key: "foreign_male", width: 10 },
        { header: "Female", key: "foreign_female", width: 10 },
        { header: "Total", key: "foreign_total", width: 10 },
        { header: "Male", key: "total_male", width: 10 },
        { header: "Female", key: "total_female", width: 10 },
        { header: "Total", key: "total_total", width: 10 },
        { header: "A", key: "age_a", width: 10 },
        { header: "B", key: "age_b", width: 10 },
        { header: "C", key: "age_c", width: 10 },
        { header: "D", key: "age_d", width: 10 },
        { header: "Please specify Country of Residence of Foreign Visitor", key: "residence", width: 30 },
    ];

    worksheet.addRow({
        name: "Amandari Cove Resort",
        municipality: "GSC",
        code: "",
        province_male: 1,
        province_female: 2,
        province_total: 3,
        other_male: 5,
        other_female: 2,
        other_total: 7,
        foreign_male: 2,
        foreign_female: 1,
        foreign_total: 3,
        total_male: 7,
        total_female: 6,
        total_total: 13,
        age_a: 1,
        age_b: 2,
        age_c: 3,
        age_d: 4,
        residence: ""
    });

    workbook.xlsx.write(unstream({}, function(buf) {
        res.status(200).send(buf);
    })).catch(err => {
        console.error(err);
    });

    /* const fileName = 'templates/DAYTIME_TOURISTS.xlsx';

    workbook.xlsx.readFile(fileName)
        .then(() => {
            let worksheet = workbook.getWorksheet("SAME DAY");

            

        }).catch(err => {
            console.log(err);
        }); */
});

http.listen(port, () => {
    console.log(`Server running or port ${port}`);
});
