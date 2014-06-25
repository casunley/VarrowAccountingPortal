/*  
 *   Use this to begin developing an understanding of the extraordinary powers of Node.js and ExpressJS!
 */
var fs = require('fs');
var express = require('express');
var path = require('path');
var app = express();
var xlsx = require('node-xlsx');
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
            //console.log(obj);
            
            for (var i=0; i < obj.worksheets.length; i++){
             
                var myObj = obj.worksheets[i];
                console.log(myObj.data);
            }
            
            //var arr = [];
            //for(worksheet in obj.worksheets)
            //{
            //    arr.push(JSON.stringify(worksheet.data));
                
          //  }
           // console.log(JSON.parse(worksheet.data));
           
        });
    });
});





