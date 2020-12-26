var express = require('express');
var app = express();
const Excel = require('exceljs');
var axios = require('axios');

app.set('view engine', 'ejs');

const workbook = new Excel.Workbook();
const worksheet = workbook.addWorksheet("My Sheet");

worksheet.columns = [
{header: 'Name', key: 'name', width: 30},
{header: 'Username', key: 'username', width: 30}, 
{header: 'Email', key: 'email', width: 30},
{header: 'ZipCode', key:'zipcode', width:15}
];


app.get('/', function(req, res){
    res.render('index.ejs');
});

app.get('/api', function(req, res){
    axios.get('https://jsonplaceholder.typicode.com/users/')
    .then(response => response.data.forEach(data => {
        worksheet.addRow({name: data.name, username: data.username, email : data.email, zipcode : data.address.zipcode});
    }));
    workbook.xlsx.writeFile('export.xlsx');
    console.log('File written');
    res.redirect('/');
});

app.get('/display', function(req, res){
    axios.get('https://jsonplaceholder.typicode.com/users/').then(response => res.render('display.ejs', {data:response.data}));
});

app.listen(4000, function(error, result){
    console.log('Server is Running!!');
});

