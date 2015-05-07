# excel-to-json
-- convert excel files to json files.
-- convert multiple nested children excel files to json files.

[![Build Status](https://travis-ci.org/vinothbabu/excel2json.svg?branch=master)](https://travis-ci.org/vinothbabu/excel2json)

## Install

```
  npm install excel-to-json
```

## Usage 1

``` javascript
  var excel2json = require("excel-to-json");
  excel2json({
    input: "input",  // input directory 
    output: "output" // output directory 
   }, function(err, result) {
    if(err) {
      console.error(err);
    } else {
      console.log(result);
    }
  });
```
## Usage 2

Alternatively, you can pass the input params by parsing a json file. 
//config.json

``` json
   {
    "name": "projectname",
    "version": "1.0.0",
    "description": "",
    "trans": {
      "input": "sample",
      "output": "output"
    }
  }
```

``` javascript
  var fs = require("fs");
  var excel2json = require("excel-to-json");
  var configuration = JSON.parse(fs.readFileSync("config.json", 'utf8'));
  excel2json(configuration, function(err, result) {
    if(err) {
      console.error(err);
    } else {
      console.log(result);
    }
  });
```
## License
wtfpl@vinoth
