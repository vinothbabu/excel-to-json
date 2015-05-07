var should = require("should");
var excel2json = require("../");
var fs = require("fs");

describe('excel to json', function(){

	it("should convert the excel file to json", function(){
		excel2json({
			"trans" : {
				"input" : "./sample",
				"output" : "./output",
			}
		},function(err, result) {
			should.not.exist(err)
			result.should.be.an.instanceOf(Object)
		})
	});

	it("should convert the nested children files to corresponding json", function(){
		excel2json({
			"trans" : {
				"input" : "./sample-children", 
				"output" : "./output"
			}
		}, function(err, result){
			should.not.exist(err);
			result.should.be.an.instanceOf(Object);
		})
	})

	it('should read file in sample-file1.json', function() {
		var exist = fs.existsSync('./output/sample/sample-file1.json')
		exist.should.be.true;
	})

});