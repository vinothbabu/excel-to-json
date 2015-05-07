var fs = require('fs');
var xlsjs = require('xlsjs');
var cvcsv = require('csv');
var async = require("async");
var mkdirp = require('mkdirp');
var path = require('path');

exports = module.exports = excel2json;

function excel2json (config) {
  if(!config.trans.input) {
    console.error("You miss a input file");
    process.exit(1);
  }

  var xj = new XlsJson();
  xj.init();
  xj.scanDirectoryTree(config.trans.input, config.trans.output);
  xj.preformXSL2JSONConversion();
}

function XlsJson(){}

XlsJson.prototype.init = function(config){
	this.files = [];	
}

XlsJson.prototype.directoryExists = function(dir) {
    try {
        return fs.statSync(dir).isDirectory();
    } catch (e) {
        return false;
    }
};
XlsJson.prototype.makeDirectory = function(dir) {
    if (this.directoryExists(dir)) return false; //skip if it exists
    try {
        fs.mkdirSync(dir);
    } catch (e) {
        if (e.code != 'EEXIST') throw e;
    }
};
XlsJson.prototype.scanDirectoryTree = function(currentDir, outputDir) {
    var that = this;
    if (!this.directoryExists(currentDir)) {
        return;
    } //ensure we have a directory
    var writeDir = path.join(outputDir, currentDir);
    this.makeDirectory(writeDir); //make the directory if it doesn't exist
    fs.readdirSync(currentDir).forEach(function(itemName) {
        var itemPath = path.join(currentDir, itemName);
        var itemInfo = fs.statSync(itemPath);
        if (itemInfo.isFile() && path.extname(itemPath).toLowerCase() == '.xls') {
            var targetName = itemName.replace(/.xls$/i, '.json');
            var readPath = itemPath;
            var writePath = path.join(writeDir, targetName);
            that.files.push({
                input: readPath,
                output: writePath
            });

        } else if (itemInfo.isDirectory()) {
            that.scanDirectoryTree(itemPath, outputDir);
        }

    });
};
XlsJson.prototype.preformXSL2JSONConversion = function() {
    function onParse(file, callback) {
        node_xj(file, callback);
    }

    function onResult(err, results) {
        if (err) throw err;
    }
    async.map(this.files, onParse, onResult); //loop through our array calling mainLogic for each item, then finally call whenFinished
};

function node_xj(config, callback) {
  var vb = new VB(config, callback);
}

function VB(config, callback) { 
  var wb = this.load_excel(config.input)
  var csv = this.getCsv(wb.Sheets[wb.SheetNames[0]]);
  this.vbjson(csv, config.output, callback);
}

VB.prototype.load_excel = function(input) {
  return xlsjs.readFile(input);
}

VB.prototype.getCsv = function(ws) {
  return csv_file = xlsjs.utils.make_csv(ws)
}

VB.prototype.vbjson = function(csv, output, callback) {
  var obj = {};
  cvcsv().from.string(csv)
    .transform( function(row){
      row.unshift(row.pop());
      return row;
    })
    .on('record', function(row, index){
        if (index > 0) obj[row[1]] = row[0];
     })     
     .on('end', function(count){
       if(output !== null) {
        var stream = fs.createWriteStream(output, { flags : 'w' });
        stream.write(JSON.stringify(obj, null, 4));
        callback(null, obj);
      } else {
        callback(null, obj);
      }
      
    })
    .on('error', function(error){
      console.log(error.message);
    });
}