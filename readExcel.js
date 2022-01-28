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
    let incompatible = ""
    // 身體質量指數 9 
    if(dataArray[9]>=35){
      incompatible = `身體質量指數(${dataArray[9]}),`
    }
    // 收縮壓 10 
    if(dataArray[10]>=160){
      incompatible = incompatible + `收縮壓(${dataArray[10]}),`
    }
    // 舒張壓 11 
    if(dataArray[11]>=100){
      incompatible = incompatible + `舒張壓(${dataArray[11]}),`
    }
    // 腰圍 12
    if(dataArray[12]>=90){
      incompatible = incompatible + `腰圍(${dataArray[12]}),`
    }
    // 尿蛋白 13
    if(dataArray[13]=="4+"){
      incompatible = incompatible + `尿蛋白(${dataArray[13]}),`
    }
    // 白血球 15
    if((dataArray[15]*1000)>=20000 || (dataArray[15]*1000)<=2500){
      incompatible = incompatible + `白血球(${dataArray[15]}),`
    }
    // 血色素 16
    if(dataArray[16]>=21 || dataArray[16]<=7){
      incompatible = incompatible + `血色素(${dataArray[16]}),`
    }
    // 丙氨酸轉氨脢 ALT 17
    if(dataArray[17]>=151){
      incompatible = incompatible + `丙氨酸轉氨脢 ALT(${dataArray[17]}),`
    }
    // 肌酸酐 18
    if(dataArray[18]>=2.5){
      incompatible = incompatible + `肌酸酐(${dataArray[18]}),`
    }
    // 總膽固醇 19
    if(dataArray[19]>=301){
      incompatible = incompatible + `總膽固醇(${dataArray[19]}),`
    }
    // 三酸甘油脂 20
    if(dataArray[20]>=501){
      incompatible = incompatible + `三酸甘油脂(${dataArray[20]}),`
    }
    // 高密度-脂蛋白 21
    if(dataArray[21]<=40){
      incompatible = incompatible + `高密度-脂蛋白(${dataArray[21]}),`
    }
    // 低密度-脂蛋白 22
    if(dataArray[22]>=191){
      incompatible = incompatible + `低密度-脂蛋白(${dataArray[22]}),`
    }
    // 空腹血糖 23
    if(dataArray[23]>=161){
      incompatible = incompatible + `空腹血糖(${dataArray[23]}),`
    }
    
    
    if(incompatible!=="") arrayList.push({
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