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

    
const jsondata = {
    "data":[
        {
        "id": "1",
        "destination" : "SG Farm",
        "municipality"  :"Gensan",
        "month":"January",
        "provinceMale":"2",
        "provinceFemale":"4",
        "otherProvinceMale":"3",
        "otherProvinceFemale":"5",
        "countAgeA":"2","countAgeB":"3",
        "countAgeC":"6","countAgeD":"8",
        "residence":[{"id":"1","country":"Ph","male":"2","female":"5","daytime_id":"1"}] 
        },{
        "id": "2",
        "destination" : "london beach",
        "municipality"  :"Gensan",
        "month":"January",
        "provinceMale":"3",
        "provinceFemale":"6",
        "otherProvinceMale":"9",
        "otherProvinceFemale":"12",
        "countAgeA":"1","countAgeB":"2",
        "countAgeC":"5","countAgeD":"7",
        "residence":[{"id":"1","country":"rh","male":"3","female":"6","daytime_id":"1"}]
        }
    
    ]
        
};

app.get("/", (req, res) => {
    res.status(201).send({access: false});
})

app.post("/report", (req, res) => {

    res.setHeader('Content-type', 'application/vnd.ms-excell');
    res.setHeader('Content-Transfer-Encoding', 'binary');
    res.setHeader('Content-disposition', 'attachment; filename="report.xlsx"');


    const fileName = 'templates/DAYTIME_TOURISTS.xlsx';
    let workbook = new excel.Workbook();

    workbook.xlsx.readFile(fileName).then(() => {

        let rowid =12;
        let worksheet = workbook.getWorksheet(1);
        let headerplacetext=jsondata.data[0].municipality.toString(); 
        let headermonthtext=jsondata.data[0].month.toString(); 
        let headerplace = worksheet.getRow(5);
        let headermonth = worksheet.getRow(6);
        headerplace.getCell(10).value=headermonthtext;
        headermonth.getCell(10).value=headerplacetext;

        for(let i =0;i < jsondata.data.length; i++){
            let row = worksheet.getRow(rowid+i);
            row.getCell(2).value = jsondata.data[i].destination.toString(); 
            row.getCell(3).value = jsondata.data[i].municipality.toString(); 
            row.getCell(5).value = parseInt( jsondata.data[i].provinceMale.toString()); 
            row.getCell(6).value = parseInt( jsondata.data[i].provinceFemale.toString()); 
            row.getCell(8).value = parseInt( jsondata.data[i].otherProvinceMale.toString());
            row.getCell(9).value = parseInt( jsondata.data[i].otherProvinceFemale.toString());
            row.getCell(18).value = parseInt( jsondata.data[i].countAgeA.toString());
            row.getCell(19).value = parseInt( jsondata.data[i].countAgeB.toString());
            row.getCell(20).value = parseInt( jsondata.data[i].countAgeC.toString());
            row.getCell(21).value = parseInt( jsondata.data[i].countAgeD.toString());
            let resData=jsondata.data[i].residence[0].country.toString() + ' male: '+ jsondata.data[i].residence[0].male.toString() + ' female: '+jsondata.data[i].residence[0].female.toString();
            row.getCell(22).value =resData;
        }
        
        workbook.xlsx.write(unstream({}, function (buf) {
            res.status(200).send(buf);
        })).catch(err => {
            console.error(err);
        });
    })
    .catch((err) => {
        console.error(err);
    });
});

http.listen(port, () => {
    console.log(`Server running or port ${port}`);
});
