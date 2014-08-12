/*  
 *   App designed to 
 */
var fs = require('fs');
var express = require('express');
var path = require('path');
var app = express();
var xlsx = require('node-xlsx');
var builder = require('xmlbuilder');
var busboy = require('connect-busboy');
var http = require('https');
var routes = require('./controllers/routes');
var posts = require('./controllers/posts/posts.js');
app.use(busboy());
app.use("/styles", express.static(__dirname + '/styles'));

var port = process.env.VCAP_APP_PORT || 3000;
app.listen(port);
console.log('Server listening on port ' + port);

app.get('/', routes.landing);
app.get('/upload', routes.upload);
app.get('/success', routes.success);
app.get('/error', routes.error);
app.get('/amexupload', routes.amexUpload);
app.get('/amexsuccess', routes.amexSuccess);


var theChunk;
var fullfile;

app.post('/upload', function (request, response) {
    var fstream;
    request.pipe(request.busboy);
    request.busboy.on('file', function (fieldname, file, filename) {
        console.log('Uploading: ' + filename);
        fstream = fs.createWriteStream('./storedFiles/' + filename);
        file.pipe(fstream);
        fstream.on('close', function () {
            response.redirect('error');
            console.log('Uploaded to ' + fstream.path);
            fullfile=path.join(__dirname, fstream.path);
            var obj = xlsx.parse(fullfile);
        
            /* 
             * Create variables for generating date (MM-DD-YYYY) in xml string
             */
            var date = new Date();
            var day = date.getDate();
            var month = date.getMonth()+1;
            var year = date.getFullYear();
            var getDateTime = month + "-" + day + "-" + year;
                        
            
            /*
             * Create XML structure using XMLBuilder lib that contains information
             * that will be pushed to Intacct
             */
            
            //First tag in XML tree
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
            var companyid = login.ele('companyid', 'Varrow-COPY');//Companyid is Varrow-COPY to access sandbox. Normally it is company id, Varrow
            var password = login.ele('password', 'V@rrowDevTeam2014');
            
            var content = operation.ele('content');
            var fnctn = content.ele('function').att('controlid', 'Varrow');
            var createbillbatch = fnctn.ele('create_billbatch');
            var batchtitle = createbillbatch.ele('batchtitle', 'Concur Batch Upload: ' + getDateTime);
            
            //For each key in
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
                            var glaccountnoStr = row[2]['value'].toString() + '-000';
                            var amountStr = row[3]['value'].toString();
                            var memoStr = row[4]['value'].toString() + ' : ' + row[6]['value'].toString();
                            var departmentidStr = row[7]['value'].toString();
                            
                            //If this is a new employee, create first set of items
                            if (empNames.indexOf(empNameStr)<0){
                                createBill = createbillbatch.ele('create_bill');
                                //Push the name onto empNames so that employee's are not duplicated
                                empNames.push(empNameStr);
                                
                                createBill.ele('vendorid', vendoridStr);
                                
                                var datecreated = createBill.ele('datecreated');
                                datecreated.ele('year', year);
                                datecreated.ele('month', month);
                                datecreated.ele('day', day);
                                
                                var datedue = createBill.ele('datedue');
                                datedue.ele('year', year);
                                datedue.ele('month', month);
                                datedue.ele('day', day);
                                
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
                   //console.log(root.toString({pretty:true}));
                   var body = root.toString({pretty:true});
            
                var len = 0
                var data = body.toString('utf8');
                len = Buffer.byteLength(data);

            
            	var req = http.request({
		          host: 'www.intacct.com',
		          path: '/ia/xml/xmlgw.phtml',
		          method: 'POST',	
		          headers: {
			     'Content-Type' : 'x-intacct-xml-request', 
			     'Content-Length' : len
		  }
	   });
            
                
                // wire up events
            
	           req.on('response', function(res){	
		          console.log('STATUS: ' + res.statusCode);
		          console.log('HTTP: ' + res.httpVersion);
		          console.log('HEADER: ' + JSON.stringify(res.headers));
		          res.setEncoding('utf8');
		          console.log('BODY (multipart):\n');
		          res.on('data', function (chunk) {
			         console.log(chunk);//Chunk is the JSON object that will be written to HTML page upon error
                    });
                   //setTimeout(function() {chunk}, 30000)//Want to hold serving redirect page until response from server has been received. 
                
  
                }).on('error', function(e) {
		          console.error(e);		
	       });
	           // send data
	           req.end(data,'utf8');


	           // output header and data read
	           console.log(req._header);	
	           console.log(data); 
            
        });
    });
}); 




/*
 * Create var called amexfullfile that holds filepath of uploaded excel file
 */



var amexfullfile;

app.post('/amexupload', function (request, response) {
    var fstream;
    request.pipe(request.busboy);
    request.busboy.on('file', function (fieldname, file, filename) {
        console.log('Uploading: ' + filename);
        fstream = fs.createWriteStream('./storedFiles/' + filename);
        file.pipe(fstream);
        fstream.on('close', function () {
            response.redirect('amexsuccess');
            console.log('Uploaded to ' + fstream.path);
            amexfullfile=path.join(__dirname, fstream.path);
            var obj = xlsx.parse(amexfullfile);
        
            /* 
             * Create variables for generating date (MM-DD-YYYY) in xml string
             */
            var date = new Date();
            var day = date.getDate();
            var month = date.getMonth()+1;
            var year = date.getFullYear();
            var getDateTime = month + "-" + day + "-" + year;
                        
            
            /*
             * Create XML structure using XMLBuilder lib that contains information
             * that will be pushed to Intacct
             */
            
            //First tag in XML tree
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
            var companyid = login.ele('companyid', 'Varrow-COPY');//Companyid is Varrow-COPY to access sandbox. Normally it is company id, Varrow
            var password = login.ele('password', 'V@rrowDevTeam2014');
            
            var content = operation.ele('content');
            var fnctn = content.ele('function').att('controlid', 'Varrow');
            var createbillbatch = fnctn.ele('create_billbatch');
            var batchtitle = createbillbatch.ele('batchtitle', 'Concur Amex Batch Upload: ' + getDateTime);
            
            //For each key in
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
                            var billnoStr = row[10]['value'].toString();
                            var descriptionStr = row[10]['value'].toString();
                            var glaccountnoStr = row[2]['value'].toString() + '-000';
                            var amountStr = row[3]['value'].toString();
                            var memoStr = row[4]['value'].toString() + ' : ' + row[6]['value'].toString();
                            var departmentidStr = row[7]['value'].toString();
                            
                            //If this is a new employee, create first set of items
                            if (empNames.indexOf(empNameStr)<0){
                                createBill = createbillbatch.ele('create_bill');
                                //Push the name onto empNames so that employee's are not duplicated
                                empNames.push(empNameStr);
                                
                                createBill.ele('vendorid', 'V-0084');
                                
                                var datecreated = createBill.ele('datecreated');
                                datecreated.ele('year', year);
                                datecreated.ele('month', month);
                                datecreated.ele('day', day);
                                
                                var datedue = createBill.ele('datedue');
                                datedue.ele('year', year);
                                datedue.ele('month', month);
                                datedue.ele('day', day);
                                
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
                   //console.log(root.toString({pretty:true}));
                   var body = root.toString({pretty:true});
            
                var len = 0
                var data = body.toString('utf8');
                len = Buffer.byteLength(data);

            
            	var req = http.request({
		          host: 'www.intacct.com',
		          path: '/ia/xml/xmlgw.phtml',
		          method: 'POST',	
		          headers: {
			     'Content-Type' : 'x-intacct-xml-request', 
			     'Content-Length' : len
		  }
	   });
            
                // wire up events
	           req.on('response', function(res){	
		          console.log('STATUS: ' + res.statusCode);
		          console.log('HTTP: ' + res.httpVersion);
		          console.log('HEADER: ' + JSON.stringify(res.headers));
		          res.setEncoding('utf8');
		          console.log('BODY (multipart):\n');
		          res.on('data', function (chunk) {
			         console.log(chunk);
                        });
                    }).on('error', function(e) {
		          console.error(e);		
	           });

	           // send data
	           req.end(data,'utf8');		

	           // output header and data read
	           console.log(req._header);	
	           console.log(data);      
        });
    });
});