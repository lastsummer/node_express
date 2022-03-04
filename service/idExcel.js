const xlsx = require('xlsx');
const crypto = require('crypto');
const XlsxPopulate = require('xlsx-populate');
const fs = require('fs')

module.exports.parserRemoveIDExcel = async function parserRemoveIDExcel(filename) {
  const excel = xlsx.readFile(filename);
  var xlData = xlsx.utils.sheet_to_json(excel.Sheets['未測量名單']);
  let arrayList = {};
  xlData.forEach((user) => {
    if(user['未測量4次原因']){
      arrayList[user['工號']] = {
        depart: user['部門名稱'],
        departNo: user['代碼'],
        two: user['上兩層組織中文名稱'],
        one: user['上一層組織中文名稱'],
        name: user['姓名'],
        reason: user['未測量4次原因']
      }
    }
  })
  fs.writeFile('removeId.json', JSON.stringify(arrayList), (error)=>{
    if(error) console.log(error)
    else console.log('success')
  })
}

module.exports.parserIDExcel = async function parserIDExcel(filename) {
  const excel = xlsx.readFile(filename);
  var xlData = xlsx.utils.sheet_to_json(excel.Sheets['工作表1']);
  let arrayList = {};
  let workArrayList = {};
  xlData.forEach((user) => {
    arrayList[user['證件號碼']] = {
      depart: user['部門名稱'],
      departNo: user['歸屬組織代碼'],
      two: user['上兩層組織中文名稱'],
      one: user['上一層組織中文名稱']
    }
    workArrayList[user['工號']] = {
      depart: user['部門名稱'],
      departNo: user['歸屬組織代碼'],
      two: user['上兩層組織中文名稱'],
      one: user['上一層組織中文名稱'],
      name: user['姓名']
    }
  })
  fs.writeFile('idTOdepart.json', JSON.stringify(arrayList), (error)=>{
    if(error) console.log(error)
    else console.log('success')
  })
  fs.writeFile('workIdTOdepart.json', JSON.stringify(workArrayList), (error)=>{
    if(error) console.log(error)
    else console.log('success')
  })
}

module.exports.parserMonthIDExcel = async function parserMonthIDExcel(filename, month) {
  const excel = xlsx.readFile(filename);
  var xlData = xlsx.utils.sheet_to_json(excel.Sheets['工作表1']);
  let workArrayList = {};
  xlData.forEach((user) => {
    workArrayList[user['工號']] = {
      depart: user['部門名稱'],
      departNo: user['歸屬組織代碼'],
      two: user['上兩層組織中文名稱'],
      one: user['上一層組織中文名稱'],
      name: user['姓名']
    }
  })
  fs.writeFile(`workIdTOdepart-${month}.json`, JSON.stringify(workArrayList), (error)=>{
    if(error) console.log(error)
    else console.log('success')
  })
}