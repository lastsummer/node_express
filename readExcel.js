const xlsx = require('xlsx')
const crypto = require("crypto")

module.exports.parserExcel = async function parserExcel(filename){
  const excel = xlsx.readFile(filename)
  var xlData = xlsx.utils.sheet_to_json(excel.Sheets["全聯實業股份有限公司總表資料"]);
  let arrayList = []
  xlData.forEach(user => {
    const dataArray = Object.keys(user).map((key)=>{
      return user[key]
    })
    let incompatible = []
    // 身體質量指數 9 
    if(dataArray[9]>=35){
      incompatible.push(`身體質量指數(${dataArray[9]}),`)
    }
    // 收縮壓 10 
    if(dataArray[10]>=160){
      incompatible.push(`收縮壓(${dataArray[10]}),`)
    }
    if(incompatible.length > 0) arrayList.push({
      "身分證字號": dataArray[6],
      "姓名": dataArray[2],
      "不符合項目":JSON.stringify(incompatible)
    })
  })
  const ws = xlsx.utils.json_to_sheet(arrayList)
  xlsx.utils.book_append_sheet(excel,ws,"4級列表")
  const id = crypto.randomBytes(20).toString('hex');
  xlsx.writeFile(excel,`result/${id}.xlsx`)
  return `${id}.xlsx`
}