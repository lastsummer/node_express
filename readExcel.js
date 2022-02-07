const xlsx = require('xlsx')
const crypto = require("crypto")

module.exports.parserHeartExcel = async function parserHeartExcel(filename){

}

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
    if(user["身體質量指數"]>=35){
      incompatible = `身體質量指數(${user["身體質量指數"]}),`
    }
    // 收縮壓 10 
    if(user["收縮壓"]>=160){
      incompatible = incompatible + `收縮壓(${user["收縮壓"]}),`
    }
    // 舒張壓 11 
    if(user["舒張壓"]>=100){
      incompatible = incompatible + `舒張壓(${user["舒張壓"]}),`
    }
    // 腰圍 12
    if(user["腰圍"]>=90){
      incompatible = incompatible + `腰圍(${user["腰圍"]}),`
    }
    // 尿蛋白 13
    if(user["尿蛋白"]=="4+"){
      incompatible = incompatible + `尿蛋白(${user["尿蛋白"]}),`
    }
    // 白血球 15
    if((user["白血球"]*1000)>=20000 || (user["白血球"]*1000)<=2500){
      incompatible = incompatible + `白血球(${user["白血球"]}),`
    }
    // 血色素 16
    if(user["血色素"]>=21 || user["血色素"]<=7){
      incompatible = incompatible + `血色素(${user["血色素"]}),`
    }
    // 丙氨酸轉氨脢 ALT 17
    if(user["丙氨酸轉氨脢 ALT"]>=151){
      incompatible = incompatible + `丙氨酸轉氨脢 ALT(${user["丙氨酸轉氨脢 ALT"]}),`
    }
    // 肌酸酐 18
    if(user["肌酸酐"]>=2.5){
      incompatible = incompatible + `肌酸酐(${user["肌酸酐"]}),`
    }
    // 總膽固醇 19
    if(user["總膽固醇"]>=301){
      incompatible = incompatible + `總膽固醇(${user["總膽固醇"]}),`
    }
    // 三酸甘油脂 20
    if(user["三酸甘油脂"]>=501){
      incompatible = incompatible + `三酸甘油脂(${user["三酸甘油脂"]}),`
    }
    // 高密度-脂蛋白 21
    if(user["高密度-脂蛋白"]<=40){
      incompatible = incompatible + `高密度-脂蛋白(${user["高密度-脂蛋白"]}),`
    }
    // 低密度-脂蛋白 22
    if(user["低密度-脂蛋白"]>=191){
      incompatible = incompatible + `低密度-脂蛋白(${user["低密度-脂蛋白"]}),`
    }
    // 空腹血糖 23
    if(user["空腹血糖"]>=161){
      incompatible = incompatible + `空腹血糖(${user["空腹血糖"]}),`
    }
    
    
    if(incompatible!=="") arrayList.push({
      "身份證字號": user["身份證字號"],
      "姓名": user["中文姓名"],
      "不符合項目":JSON.stringify(incompatible)
    })
  })
  const ws = xlsx.utils.json_to_sheet(arrayList)
  xlsx.utils.book_append_sheet(excel,ws,"4級列表")
  const id = crypto.randomBytes(20).toString('hex');
  xlsx.writeFile(excel,`result/${id}.xlsx`)
  return `${id}.xlsx`
}