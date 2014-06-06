/*  
*    Super Easy File Uploader
*   @author greed@varrow.com
*   June 2014
*   Use this to begin developing an understanding of the extraordinary powers of Node.js and ExpressJS!
*/ 
var fs = require('fs');
var express = require('express');
var app = express();
var busboy = require('connect-busboy');
app.use(busboy());

var port = 3000;
app.listen(port);
console.log('Server listening on port ' + port);

app.get('/', function(request, response) {
	fs.readFile('upload.html', function (error, data) {
		if(error) { 
            console.log(error); 
        }
	    response.writeHead(200, {'Content-Type': 'text/html', 'Content-Length': data.length});
	    response.write(data);
    	response.end();
	});
});

app.post('/upload',function(request, response) {
	var fstream;
    request.pipe(request.busboy);
    request.busboy.on('file', function(fieldname, file, filename) {
    	console.log('Uploading: ' + filename);
    	fstream = fs.createWriteStream('./tehFiles/' + filename);
    	file.pipe(fstream);
    	fstream.on('close', function() {
    		response.redirect('back');
            console.log('Uploaded to ' + fstream.path);
    	});
    });
});