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
app.get('/amexupload', routes.amexUpload);
app.get('/amexsuccess', routes.amexSuccess);
app.post('/upload', posts.upload);
app.post('/amexupload', posts.amexUpload);