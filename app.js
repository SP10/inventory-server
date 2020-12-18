var express = require('express');
var app = express();
var excel = require('./excel');

app.get('/', function (req, res) {
  res.send('Server are work!');
  excel.init();
});

app.listen(9000, function () {
  console.log('Example app listening on port 9000!');
});