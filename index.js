var MongoClient = require('mongodb').MongoClient;
var xlsx = require("xlsx");
var csv = require('csvjson');
//default url untuk mongoDB
var url = "mongodb://localhost:27017/";

// untuk buka workbook
var wb = xlsx.readFile("sample2.xlsx");

// untuk buka worksheet
var ws = wb.Sheets["Orders"];

// merubah data jadi csv
var data = xlsx.utils.sheet_to_csv(ws);

// untuk menghilangkan tanda kutip(")
var rmv = data.replace(/"/g,"");

//mengubah bentuk dari csv ke object
var dataObj = csv.toObject(rmv);

//masukkan data ke mongodb
MongoClient.connect(url, function(err, db) {
  if (err) throw err;
  //nama databasenya
  var dbo = db.db("superstore");
  //nama collectionnya
  dbo.collection("orders").insert(dataObj, function(err, res) {
    if (err) throw err;
    console.log("Document inserted");
    db.close();
  });
});