/*
 * Routes to retrieve each of the html web pages
 * app.get functions reside in app.js
 */

var fs = require('fs'); 

exports.landing = function (request, response) {
    fs.readFile('public/views/landing.html', function (error, data) {
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
}


//GET function used to point router to retrieve upload.html page and display it
exports.upload = function (request, response) {
    fs.readFile('public/views/upload.html', function (error, data) {
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
}

//GET function used to point router to retrieve success.html page. 
exports.success = function (request, response) {
    fs.readFile('public/views/success.html', function (error, data) {
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
}

//GET function used to point router to retrieve amexupload.html page
exports.amexUpload = function (request, response) {
    fs.readFile('public/views/amexupload.html', function (error, data) {
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
}

//GET function used to point router to retrieve amexsuccess.html page
exports.amexSuccess = function (request, response) {
    fs.readFile('public/views/amexsuccess.html', function (error, data) {
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
}