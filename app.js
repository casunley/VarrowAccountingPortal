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
            console.log(getDateTime);
            
                        
                                                                           
            //Will print out parsed excel sheet data
            for (var i=0; i < obj.worksheets.length; i++)
            {             
                var myObj = obj.worksheets[i];
                console.log(myObj.data);               
              }
            
            /*
             * Creates static portion of XML sheet that will be pushed to Intacct and prints it
             * to the console
             */
            var doc = builder.create('request', {'version': '1.0', 'encoding': 'UTF-8'});                   
            var staticXML = {
                '#list': [
                            {
                                control : {
                                    'senderid': 'Varrow',
                                    'password': 'KYloh3jU0W',
                                    'controlid': 'Varrow',
                                    'dtdversion': '2.1'
                                        },
            
                                operation: {
                                    authentication: {
                                        login: {
                                            'userid': 'dsgroup',
                                            'companyid': 'Varrow',
                                            'password': 'V@rrowDevTeam2014'
                                                }
                                            },
                                    content: {
                                        'function': {'controlid': 'Varrow'},
                                        create_billbatch: {
                                            'batchtitle': 'Conucur Batch Upload: ' + getDateTime
                                     }
                                  }
                                }
                            }
                        ]
                    };
            
            doc.ele(staticXML);
            console.log(doc.toString({ pretty: true }));
        
    
        });
    });
});





