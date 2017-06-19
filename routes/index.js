var express = require('express');
var router = express.Router();
var GoogleSpreadsheet = require('google-spreadsheet');
var async = require('async');
var JSONSecure = {
  "type": "service_account",
  "project_id": "crisheets-171201",
  "private_key_id": "541e3c34ad3633d40654385d809aa40bdfdc7b72",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDNgGGnIqZfONqG\nguiiNAB4fnY8ag2unSpRMKkdOOB1IcqWzOe0DI0DlvDfY5wBYRajMU3kwj5czId4\nOaq7v+rbRvH4wyLqTAlUyingxASMUdkJNDXCydUnvSmIwS7erkSvVmxecCdDfcIf\ncXVXtEx9OC/D+zTa8HdSzOOG8CifVngwFEEikPFexuhxY0isWydBo99dqMKVs4HS\nEI9WXbkuQs5JDtNjAuqGkd92hP3gQ4KeI3pmx9j26PRXr3cuJlkqFRoMBtheDcvb\nhY9KSLAHF7+tYMHLiWNDbnY6mHHLBB/6kfIJHi8hBDCHzPIV1HTxrZi4DHySRWY0\noNYEJnhpAgMBAAECggEADtGxORJAmSFKcOTDzd6eIhJMiHLFZdAjmxM9PsQ5O3ki\nWNmR4+P7z0R1Pka6m0bbEQ2fJl3zjVGae0r9Sui+EO3+yQeZXvf4vEqP1ouMIGpO\nkmFXdyKC4oi7lAcKUDiU72f2DKFd03ogI5BU++wej/EbULCu0RMCJ9Bqn/J1m3fJ\npezgbvt925790O5DK2aTDiZ4lSNMTaIXli7y1CO58aclhSv9TC9JuRSbFxvrwVz8\n048JCub8HuoDMrXkbJJL0I7MkPZUnI9XXiAdzxqA+RUzF5EMYga+IqjUMtPBayCV\nBulNXBfP+od2Ww7F4V1hUKiHqiPkLnqBspmLuxp/uQKBgQD49g2c5YUZusWCFHqI\nHlCq7nm8o3TuA1SBVSKFg+32neB9cNrOczj7ZarBAJEgyFQ2fxNpmyGu3OYsuA+N\nEFCmjuQ4bbjPvJ//4hY4fWmYVA1KZ7oRZaQQiDK2c05W0KYf28/TmKaSmZ3ihjvF\n9AlD/a4pYnu0tyvVnFvrdLIQHwKBgQDTT8Xx7HKbx0b7jAgOW+UVsMsVo8h/sYz6\nn3XWhQ3RqAz9u6D7EY7lP/pJDXYJqKHvPz8NHYzd1oKAu2h+MBrLsbEXXcu6QTpy\n80mkMe7MGkbK0RK0IbDhHk80CuTny5DNVmzzkhi0WDHuGY6FgKDnLHMfBTQin0K5\nKUPQoffGdwKBgQCxeMVPeRYet1OgXPTUH7glgYZqgKMUIG+XGpdXUirKSNUE1vRD\ng0O1gk3s83iBRA00I/Y0rA+g1XuHmVYMmvrDIM1zpFPAUphEEmkAr/YcTp6C0dqE\nGE6SMmTkRuIrZOVnhIxsLD9h9fvkxQfLHyGTxDzo53mD46dtyN6FxFRCxQKBgE0e\nehF2x8UOj1tSmcYTx6GI6jU1lwDzXY5CEBGAcbCockVP6sp2d/42wTNUFFYmEyNr\no6k3tadomCw/OT2EdOMOMKFke+u3zosROzkPeCVJGbj/YKIZAaLHGwTVTQFDi89E\n0XJ7SUTYQzCwZxFWBmMYF9OkZIiWMxW8d6F22yS1AoGBAO4Z0pqyPLHpnGJ8J89I\njmmrQM4ZmqLS6FJs7TArXxikQMFAK42EwLgwRjV0RJ+YeC4v6HZEgQPiKDx0RfLL\nvUEVgdhilbRHuwQ569TCAHnR+Lzw5XH9/+Q09JBWOHwh2VT1zc9uk2G7/dscQytg\ntdexmeKq4emVbBqVcHapoDE5\n-----END PRIVATE KEY-----\n",
  "client_email": "nickjanne@crisheets-171201.iam.gserviceaccount.com",
  "client_id": "113793794730290092629",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://accounts.google.com/o/oauth2/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/nickjanne%40crisheets-171201.iam.gserviceaccount.com"
}
var maxRows;
var maxCols;
/* GET home page. */
router.get('/', function(req, res, next) {
  res.render('index');
});

module.exports = router;


router.get('/test', function(req, res, next) {
var doc = new GoogleSpreadsheet('12_A4j3ii-H5m0QpqmknpFAxyUre7FDnKrYwUpeuB-J8');
var sheet;

async.series([
  function setAuth(step) {
    var creds = JSONSecure;

    doc.useServiceAccountAuth(creds, step);
  },
  function getInfoAndWorksheets (step) {
    doc.getInfo(function(err, info) {
      console.log("Loaded doc: " +info.title+' by '+info.author.email);
      sheet = info.worksheets[0];
      console.log("Sheet 1:"+sheet.title+' '+sheet.rowCount+'x'+sheet.colCount);
      maxCols = sheet.colCount;
      step();
    });
  },
  function workingWithRows (step) {
    sheet.getRows({
      offset: 1,
      limit: 60,
      orderby: 'col2'
    }, function(err, rows) {
      console.log("Read "+rows.length+" rows");

      maxRows = rows.length+1;
      step();
    });
  },
  function workingWithCells(step) {
    sheet.getCells({
      'min-row': 4,
      'max-row': maxRows,
      'min-col': 2,
      'max-col': 2,
      'return-empty': true
    }, function(err, cells) {

      var CellData = {};
      //Get Rowers
      var rowers = {};
      for (var i = 0; i<cells.length; i++) {
        rowers[i] = cells[i].value;
      }
      //Get Who is coming

      var date = "6/18" //NEED TO CHANGE DATE TO REAL DATE



      var dateOptions = {
        'min-row': 2,
        'max-row': 2,
        'return-empty': true
      };

      var finalComing = {};
      var dateColumn = "";

      sheet.getCells(dateOptions, function(err, cells) {
        for (var i = 0; i<cells.length; i++) {
          if (date == cells[i].value){
          dateColumn = cells[i].col;
          console.log("DateCol " + dateColumn);
          break;
          }
        }
      });

      var columnOptions = {
        'min-row': 4,
        'max-row': maxRows,
        'min-col': dateColumn,
        'max-col': dateColumn,
        'return-empty': true
      };


      sheet.getCells(columnOptions, function(err, cells) {
        for (var i = 0; i<cells.length; i++) {
          var name = rowers[i.toString()];
          if (cells[i].value === "1") {
            finalComing[name] = "Coming for sure";
          } else if (cells[i].value === "0") {
            finalComing[name] = "Not coming for sure";
          } else {
            finalComing[name] = "Did not fill out spreadsheet/correctly";
          }
        }
        res.json(finalComing);
      });



      step();

    });
  }
], function(err) {
  if (err) {
    console.log("Error: "+err);
  }
});
});
