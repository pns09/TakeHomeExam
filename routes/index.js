var express = require('express')
var router = express.Router()
var _ = require('underscore')
var Excel = require('exceljs')
var fs = require('fs')
var config = require('../config/config')
var xlsxj = require("xlsx-to-json");

/* GET home page. */
router.get('/', function(req, res, next) {
  res.render('index', { title: 'Express' })
})

router.post('/',(req,res,next) => {
  console.log(req.file)
  xlsxj({
    input: './uploads/input.xlsx', 
    output: null
  }, function(err, jsondata) {
    
    if(err) {
      console.error(err)
    }else {
      console.log(jsondata)
      var sno = 1
      var sortedjsondata = _.sortBy(jsondata, function(ele){
        ele['Credit Score'] = ele['Critic Score']
        delete ele['Critic Score']
        return -ele['Credit Score'] && ele.Genre
      })
      _.each(sortedjsondata,function(ele){
        ele.SNO = sno
        sno = sno + 1
      })
      var workbook = new Excel.Workbook();
      var sheet = workbook.addWorksheet('Output')
      sheet.font = config.font
      _.each(config.author.label,function(value,key,list){
          sheet.getCell('B2')[key] = value
      })
      
      sheet.mergeCells('C2:D2')
      _.each(config.author.signature, function(value,key,list){
          sheet.getCell('D2')[key] = value
      })
      var cellcolumn = config.columns
      var headers = config.headervalues
      for (var i = 0;i < 6;i++) {
          sheet.getCell(cellcolumn[i]+4).value = headers[i]
          sheet.getColumn(i+1).width = 20
          _.each(config.header, function (value, key, list) {
              sheet.getCell(cellcolumn[i] + 4)[key] = value
          })
      }
      var changecolor = 'FFFFF2CC'
      var changecolor1 = 'FFC6E0B4'
      var genre = sortedjsondata[0]['Genre']
      for(var i = 0; i < sortedjsondata.length;i++){
          if (genre != sortedjsondata[i]['Genre']) {
              genre = sortedjsondata[i]['Genre']
                  temp = changecolor
                  changecolor = changecolor1
                  changecolor1 = temp
          }
          for(var j = 0; j < 6; j++){
              var row = i + 5
              sheet.getCell(cellcolumn[j] + row).value = sortedjsondata[i][headers[j]]
              sheet.getCell(cellcolumn[j] + row).fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: changecolor }
              }
              sheet.getCell(cellcolumn[j] + row).border = config.border
              sheet.getCell(cellcolumn[j] + row).alignment = {
                  horizontal: /\d/.test(sortedjsondata[i][headers[j]]) ? 'right' : 'left'
              }
          }        
      }
      workbook.xlsx.writeFile('./public/' + config.outputfilename).then(function () {
         res.send('download ready')
      })
     .catch(err => {
       res.send(err)
     })   
    
    }
  })
})

module.exports = router
