// External Requirements
var express = require('express');
var path = require('path');
var bodyParser = require('body-parser');
var jsforce = require('jsforce');
var url = require('url');
var fs = require('fs');
var xlsx = require('node-xlsx');
var builder = require('xmlbuilder');
var http = require('https');
var jade = require('jade');
var xmlToJs = require('xml2js').parseString;
var busboy = require('connect-busboy');

// Create the app
var app = express();
app.use(busboy());

// Jade Engine setup
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'jade');
app.use(bodyParser.json());
app.use(bodyParser.urlencoded());
app.use(express.static(path.join(__dirname, 'public')));

// Set up the Oauth2 Info for jsforce
var oauth2 = new jsforce.OAuth2({
  clientId: '3MVG9VmVOCGHKYBSCAfWkFveQdwU4SOrIxxOMKuXxRGzaGOGgkkPBkazLRnoLZWf0NzbUEptPRzEbXlMydf2g',
  clientSecret: '2686887070428587222',
  redirectUri: 'http://localhost:8080/_auth'
});

// Create the connection using the Oauth2 Info
var conn = new jsforce.Connection({
  oauth2: oauth2
});

// Base URL redirects to authentication if not authorized
// Otherwise, redirect to search page
app.get('/', function(req, res) {
   if (!conn.accessToken) {
    res.redirect(oauth2.getAuthorizationUrl({
      scope: 'api id web'
    }));
  }else {
    res.redirect('/landing');
  }
});

// Using jsforce, this redirects to Salesforce Oauth2 Page
// Gets authorization code and authenticates the user
app.get('/_auth', function(req, res) {
  var code = req.param('code');
  conn.authorize(code, function(err, userInfo) {
    if (err) {
      return console.error(err);
      res.render('index', {
        title: 'Could Not Connect!'
      });
    } else {
      var user;
      var queryString = ('SELECT Id, Name FROM User WHERE ' +
        'Id = \'' + userInfo.id + '\' LIMIT 1');
      conn.query(queryString)
        .on("record", function(record) {
          user = record;
        })
        .on("end", function(query) {
          app.locals.auth = conn.accessToken;
          app.locals.user = user.Name;
          app.locals.userId = user.Id;
          res.redirect('/landing');
        })
        .on("error", function(err) {
          console.error(err);
        })
        .run({
          autoFetch: true,
          maxFetch: 4000
        });
    }
  });
});

app.get('/search', function(req, res) {
  if (!conn.accessToken) {
    res.redirect(oauth2.getAuthorizationUrl({
      scope: 'api id web'
    }));
  } else {
    res.render('search', {
      title: 'Search for a Varrow SFDC account',
    });
  }
  
});

app.get('/landing', function(request, response) {
  response.render('landing', {
    title: 'Accounting Landing Page',
  });
});

app.get('/upload', function(request, response) {
  response.render('upload', {
    title: 'Upload Employee Expenses',
  });
});

app.get('/amexupload', function(request, response) {
  response.render('amexupload', {
    title: 'Upload Amex Expenses',
  });
});

//fullfile stores the path where the excel sheets are stored after downloading
var fullfile;

app.post('/upload', function(request, response) {
  var fstream;
  request.pipe(request.busboy);
  request.busboy.on('file', function(fieldname, file, filename) {
    console.log('Uploading: ' + filename);
    fstream = fs.createWriteStream('./storedFiles/' + filename);
    file.pipe(fstream);
    fstream.on('close', function() {
      console.log('Uploaded to ' + fstream.path);
      fullfile = path.join(__dirname, fstream.path);
      var obj = xlsx.parse(fullfile);

      /* 
       * Create variables for generating date (MM-DD-YYYY) in xml string
       */
      var date = new Date();
      var day = date.getDate();
      var month = date.getMonth() + 1;
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
      var companyid = login.ele('companyid', 'Varrow-COPY'); //Companyid is Varrow-COPY to access sandbox. Normally it is company id, Varrow
      var password = login.ele('password', 'V@rrowDevTeam2014');

      var content = operation.ele('content');
      var fnctn = content.ele('function').att('controlid', 'Varrow');
      var createbillbatch = fnctn.ele('create_billbatch');
      var batchtitle = createbillbatch.ele('batchtitle', 'Concur Batch Upload: ' + getDateTime);

      //For each key in
      for (var key in obj.worksheets[0]) {
        //Get the value of the worksheet
        if (obj.worksheets[0].hasOwnProperty(key)) {
          var val = obj.worksheets[0][key];
          //Employee list for tracking duplicates
          var empNames = [];
          //Start loop at one to skip first row of headers
          for (var i = 1; i < val.length; i++) {
            //Row objects start here
            var row = val[i];
            var createBill;

            //Check to see if empty object from excel parser
            if (row[0]['value'] != undefined) {
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
              if (empNames.indexOf(empNameStr) < 0) {
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
      var body = root.toString({
        pretty: true
      });

      var len = 0
      var data = body.toString('utf8');
      len = Buffer.byteLength(data);

      var req = http.request({
        host: 'www.intacct.com',
        path: '/ia/xml/xmlgw.phtml',
        method: 'POST',
        headers: {
          'Content-Type': 'x-intacct-xml-request',
          'Content-Length': len
        }
      });
      // wire up events            
      req.on('response', function(res) {
        console.log('STATUS: ' + res.statusCode);
        console.log('HTTP: ' + res.httpVersion);
        console.log('HEADER: ' + JSON.stringify(res.headers));
        res.setEncoding('utf8');
        console.log('BODY (multipart):\n');
        res.on('data', function(chunk) {
          console.log(chunk); //Chunk is the JSON object that will be written to HTML page upon error
          // Convert the chunk xml string to json using xml2js
          // Then stringify the result and parse it back into json to sanatize
          xmlToJs(chunk, function(err, result) {
            chunkJson = JSON.stringify(result);
            //console.log(chunkJson);
            var parsedEmps = JSON.parse(JSON.stringify(empNames));
            console.log(parsedEmps);
            var chunkParsedJson = JSON.parse(chunkJson);
            uploadresult = JSON.stringify(chunkParsedJson['response']['operation'][0]['result'][0]['status'][0]);
            if (uploadresult == '\"failure\"') {
              response.render('errorpage', {
                title: 'Error!',
                chunk: chunk
              });
            } else {
              response.render('success', {
                title: 'Successful Upload',
                chunk: chunk,
                empNames: empNames  
              })
            }
          });
        });
      }).on('error', function(e) {
        console.error(e);
        theChunk += e;
        success = false;
      });
      // send data
      req.end(data, 'utf8');
      //output header and data read
      console.log(req._header);
      console.log(data);
    });
  });
});

//Create var to hold filepath of uploaded excel sheet
var amexfullfile;

app.post('/amexupload', function(request, response) {
  var fstream;
  request.pipe(request.busboy);
  request.busboy.on('file', function(fieldname, file, filename) {
    console.log('Uploading: ' + filename);
    fstream = fs.createWriteStream('./storedFiles/' + filename);
    file.pipe(fstream);
    fstream.on('close', function() {
      console.log('Uploaded to ' + fstream.path);
      amexfullfile = path.join(__dirname, fstream.path);
      var obj = xlsx.parse(amexfullfile);

      /* 
       * Create variables for generating date (MM-DD-YYYY) in xml string
       */
      var date = new Date();
      var day = date.getDate();
      var month = date.getMonth() + 1;
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
      var companyid = login.ele('companyid', 'Varrow-COPY'); //Companyid is Varrow-COPY to access sandbox. Normally it is company id, Varrow
      var password = login.ele('password', 'V@rrowDevTeam2014');

      var content = operation.ele('content');
      var fnctn = content.ele('function').att('controlid', 'Varrow');
      var createbillbatch = fnctn.ele('create_billbatch');
      var batchtitle = createbillbatch.ele('batchtitle', 'Concur Amex Batch Upload: ' + getDateTime);

      //For each key in
      for (var key in obj.worksheets[0]) {
        //Get the value of the worksheet
        if (obj.worksheets[0].hasOwnProperty(key)) {
          var val = obj.worksheets[0][key];
          //Employee list for tracking duplicates
          var empNames = [];
          //Start loop at one to skip first row of headers
          for (var i = 1; i < val.length; i++) {
            //Row objects start here
            var row = val[i];
            var createBill;

            //Check to see if empty object from excel parser
            if (row[0]['value'] != undefined) {
              //Items we need from excel parser
              var empNameStr = row[0]['value'].toString();
              var billnoStr = row[10]['value'].toString();
              var descriptionStr = row[10]['value'].toString();
              var glaccountnoStr = row[2]['value'].toString() + '-000';
              var amountStr = row[3]['value'].toString();
              var memoStr = row[4]['value'].toString() + ' : ' + row[6]['value'].toString();
              var departmentidStr = row[7]['value'].toString();

              //If this is a new employee, create first set of items
              if (empNames.indexOf(empNameStr) < 0) {
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
      var body = root.toString({
        pretty: true
      });

      var len = 0
      var data = body.toString('utf8');
      len = Buffer.byteLength(data);


      var req = http.request({
        host: 'www.intacct.com',
        path: '/ia/xml/xmlgw.phtml',
        method: 'POST',
        headers: {
          'Content-Type': 'x-intacct-xml-request',
          'Content-Length': len
        }
      });

      // wire up events            
      req.on('response', function(res) {
        console.log('STATUS: ' + res.statusCode);
        console.log('HTTP: ' + res.httpVersion);
        console.log('HEADER: ' + JSON.stringify(res.headers));
        res.setEncoding('utf8');
        console.log('BODY (multipart):\n');
        res.on('data', function(chunk) {
          console.log(chunk); //Chunk is the JSON object that will be written to HTML page upon error
          // Convert the chunk xml string to json using xml2js
          // Then stringify the result and parse it back into json to sanatize
          xmlToJs(chunk, function(err, result) {
            chunkJson = JSON.stringify(result);
            //console.log(chunkJson);
            var chunkParsedJson = JSON.parse(chunkJson);
            uploadresult = JSON.stringify(chunkParsedJson['response']['operation'][0]['result'][0]['status'][0]);
            if (uploadresult == '\"failure\"') {
              response.render('errorpage', {
                title: 'Error!',
                chunk: chunk
              });
            } else {
              response.render('amexsuccess', {
                title: 'Successful Upload',
                chunk: chunk
              })
            }
          });
        });
      }).on('error', function(e) {
        console.error(e);
        theChunk += e;
        success = false;
      });
      // send data
      req.end(data, 'utf8');
      //output header and data read
      console.log(req._header);
      console.log(data);
    });
  });
});

// Search results
app.post('/search/results', function(req, res) {
  var searchKey = req.param('searchKey');
  console.log(searchKey);
  var records = [];
  var queryString = ('SELECT Id, Name FROM Account WHERE ' +
    '(Name LIKE\'\%' + searchKey + '\%\' OR Name ' +
    'LIKE\'\%' + searchKey + '\' OR Name LIKE \'' +
    searchKey + '\%\') AND (NOT Name LIKE \'\%Leasing\%\')');
  conn.query(queryString)
    .on("record", function(record) {
      records.push(record);
    })
    .on("end", function(query) {
      console.log("total fetched : " + query.totalFetched);
      res.render('results', {
        title: query.totalFetched + ' Account Record(s) found for search: ' + searchKey,
        records: records
      });
    })
    .on("error", function(err) {
      console.error(err);
    })
    .run({
      autoFetch: true,
      maxFetch: 4000
    });
});

// Account Detail Page - Lots of Queries.
app.get('/account-detail/:id', function(req, res) {
  var id = req.param('id');
  var account;
  var salesReps = [];
  var queryString1 = ('SELECT Id, Name, SalesRep__c FROM Account WHERE ' +
    'Id = \'' + id + '\' LIMIT 1');
  var queryString2 = ('select Id, Name, QuickBook_Initals__c, IsActive FROM User where QuickBook_Initals__c != \'\' and QuickBook_Initals__c != \'CJJ\' and IsActive = True order by Name');
  conn.query(queryString1)
    .on("record", function(record) {
      account = record;
    })
    .on("end", function(query) {
      console.log("Fetched: " + account.Name);
      conn.query(queryString2)
        .on("record", function(record) {
          salesReps.push(record);
        })
        .on("end", function(query) {
          console.log("Fetched: " + account.Name);
          res.render('account-detail', {
            title: 'Account Records',
            account: account,
            salesReps: salesReps
          });
        })
        .on("error", function(err) {
          console.error(err);
        })
        .run({
          autoFetch: true,
          maxFetch: 4000
        });
    })
    .on("error", function(err) {
      console.error(err);
    })
    .run({
      autoFetch: true,
      maxFetch: 4000
    });
});

// Logout revokes the access token from the server and client
app.get('/logout', function(req, res) {
  conn.logout();
  conn.accessToken = '';
  res.render('index', {
    title: 'Disconnected from Salesforce!',
    auth: null
  });
});

// Catch 404 and forward to error handler
app.use(function(req, res, next) {
  var err = new Error('Not Found');
  err.status = 404;
  next(err);
});

// Start Listening!
app.listen(8080);
console.log('Express Server listening on port 8080');