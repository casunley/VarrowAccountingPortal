/*  
 *   Use this to begin developing an understanding of the extraordinary powers of Node.js and ExpressJS!
 */
var fs = require('fs');
var express = require('express');
var path = require('path');
var app = express();
var xlsx = require('node-xlsx');
var builder = require('xmlbuilder');
var busboy = require('connect-busboy');
app.use(busboy());
app.use("/styles", express.static(__dirname + '/styles'));

var port = 3000;
app.listen(port);
console.log('Server listening on port ' + port);

/*
 * GET function used to point router to retrieve upload.html page and display it
 */
app.get('/', function (request, response) {
    fs.readFile('public/upload.html', function (error, data) {
        if (error) {
            console.log(error);
        }
        response.writeHead(200, {
            'Content-Type': 'text/html',
            'Content-Length': data.length
        });
        response.write(data);
        response.end();
    });
});

/*
 *  GET function used to point router to retrieve success.html page. 
 */
app.get('/success', function (request, response) {
    fs.readFile('public/success.html', function (error, data) {
        if (error) {
            console.log(error);
        }
        response.writeHead(200, {
            'Content-Type': 'text/html',
            'Content-Length': data.length
        });
        response.write(data);
        response.end();
    });
});

/*
 * Create var called fullfile that holds filepath of uploaded excel file
 * HTTP post request
 */
var fullfile;

app.post('/upload', function (request, response) {
    var fstream;
    request.pipe(request.busboy);
    request.busboy.on('file', function (fieldname, file, filename) {
        console.log('Uploading: ' + filename);
        fstream = fs.createWriteStream('./storedFiles/' + filename);
        file.pipe(fstream);
        fstream.on('close', function () {
            response.redirect('success');
            console.log('Uploaded to ' + fstream.path);
            fullfile=path.join(__dirname, fstream.path);
            var obj = xlsx.parse(fullfile);
        
            /* 
             * Create variables for generating date (MM-DD-YYYY) in xml string
             * 
             */
            var date = new Date();
            var day = date.getDate();
            var month = date.getMonth()+1;
            var year = date.getFullYear();
            var getDateTime = month + "-" + day + "-" + year;
                        
            
            var root = builder.create('request');
            
            var control = root.ele('control');
            var senderid = control.ele('senderid', 'Varrow');
            var password = control.ele('password', 'KYloh3jU0W');
            var controlid = control.ele('controlid', 'Varrow');
            var dtdversion = control.ele('dtdversion', '2.1');
            
            var operation = root.ele('operation');
            var authentication = operation.ele('authentication');
            var login = authentication.ele('login');
            var userid = login.ele('userid', 'dsgroup');
            var companyid = login.ele('companyid', 'Varrow');
            var password = login.ele('password', 'V@rrowDevTeam2014');
            
            var content = operation.ele('content');
            var fnctn = content.ele('function').att('controlid', 'Varrow');
            var createbillbatch = fnctn.ele('create_billbatch');
            var batchtitle = createbillbatch.ele('Concur Batch Upload: ' + getDateTime);
            
            for (var key in obj.worksheets[0]) {
                 //Get the value of the worksheet
                 if (obj.worksheets[0].hasOwnProperty(key)){
                    var val = obj.worksheets[0][key];
                    //Employee list for tracking duplicates
                    var empNames = [];
                     //Start loop at one to skip first row of headers
                    for (var i = 1; i < val.length; i++){
                        //Row objects start here
                        var row = val[i];
                        var createBill;
                        
                        //Check to see if empty object from excel parser
                        if(row[0]['value'] != undefined){
                            //Items we need from excel parser
                            var empNameStr = row[0]['value'].toString();
                            var vendoridStr = row[1]['value'].toString();
                            var billnoStr = row[10]['value'].toString();
                            var descriptionStr = row[10]['value'].toString();
                            var glaccountnoStr = row[2]['value'].toString(); + '-000';
                            var amountStr = row[3]['value'].toString();
                            var memoStr = row[4]['value'].toString() + ' : ' + row[6]['value'].toString();
                            var departmentidStr = row[7]['value'].toString();
                            
                            //If this is a new employee, create first set of items
                            if (empNames.indexOf(empNameStr)<0){
                                createBill = batchtitle.ele('create_bill');
                                //Push the name onto empNames so that employee's are not duplicated
                                empNames.push(empNameStr);
                                
                                createBill.ele('vendorid', vendoridStr);
                                
                                var datecreated = createBill.ele('datecreated');
                                datecreated.ele('year', year);
                                datecreated.ele('month', month);
                                datecreated.ele('day', day);
                                
                                createBill.ele('billno', billnoStr);
                                createBill.ele('description', descriptionStr);
                                
                                var billItems = createBill.ele('billitems');
                                var lineItem = billItems.ele('lineitem');
                                lineItem.ele('glaccountno', glaccountnoStr);
                                lineItem.ele('amount', amountStr);
                                lineItem.ele('memo', memoStr);
                                lineItem.ele('departmentid', departmentidStr);
                            }
                            //If employee exists, only add additional line items
                            else {
                                var lineItem = billItems.ele('lineitem');
                                lineItem.ele('glaccountno', glaccountnoStr);
                                lineItem.ele('amount', amountStr);
                                lineItem.ele('memo', memoStr);
                                lineItem.ele('departmentid', departmentidStr);
                            }
                        } else {
                            //Row did not contain valid data.
                        }
                    }
                }
            }
                   console.log(root.toString({pretty:true}));
        });
    });
});
            

            
                               
            




