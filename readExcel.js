const xlsx = require('xlsx')
const crypto = require("crypto")

function formatDate(numb){
  if(numb!= undefined){
    let time = new Date((numb-1) * 24 * 3600000 + 1)
    time.setYear(time.getFullYear()-70)
    return time.getFullYear()
  }
}

function validateIdNumberToAgeYear(str){
  let date = new Date();
  let currentYear = date.getFullYear();
  let birdthdayArr = str.split("-");
  let year = birdthdayArr[0]
  if(birdthdayArr.length==1){
    birdthdayArr = str.split("/");
    year = birdthdayArr[0]
  }
  if(birdthdayArr.length==1){
    year = formatDate(str)
  }
  return currentYear - year

}
async function parserHeartExcel(filename){
  const excel = xlsx.readFile(filename)
  var xlData = xlsx.utils.sheet_to_json(excel.Sheets["全聯實業股份有限公司總表資料"]);
  let arrayList = []
  xlData.forEach(user => {
    let totaScore = 0;
    let ageScore = 0;
    let cholesterolScore = 0;
    let highCholesterolScore = 0;
    let pressScore = 0;
    let sickScore = 0;
    let smokeScore = 0;
    let probability = "";
    const age = validateIdNumberToAgeYear(user["出生日期"]+"");
    const sex = user["性別"];
    const cholesterol = user["總膽固醇"];
    const highCholesterol = user["高密度-脂蛋白"];
    const sbp = user["收縮壓"]
    const dbp = user["舒張壓"]
    const sick = user["慢性病史"]
    const smoke = user["抽菸"]
    if(age!="" && age>=30){
      if(sex=="男"){
        if(age<=34) ageScore = -1;
        else if(age<=39) ageScore = 0;
        else if(age<=44) ageScore = 1;
        else if(age<=49) ageScore = 2;
        else if(age<=54) ageScore = 3;
        else if(age<=59) ageScore = 4;
        else if(age<=64) ageScore = 5;
        else if(age<=69) ageScore = 6;
        else if(age<=74) ageScore = 7;

        if(cholesterol<160) cholesterolScore = -3;
        else if(cholesterol<=199) cholesterolScore = 0;
        else if(cholesterol<=239) cholesterolScore = 1;
        else if(cholesterol<=279) cholesterolScore = 2;
        else if(cholesterol>=280) cholesterolScore = 3;

        if(highCholesterol<35) highCholesterolScore = 2;
        else if(highCholesterol<=44) highCholesterolScore = 1;
        else if(highCholesterol<=49) highCholesterolScore = 0;
        else if(highCholesterol<=59) highCholesterolScore = 0;
        else if(highCholesterol>=60) highCholesterolScore = -2;

        if(sbp<120 && dbp<80) pressScore = 0;
        else if(sbp<129 || dbp<84) pressScore = 0;
        else if(sbp<139 || dbp<89) pressScore = 1;
        else if(sbp<159 || dbp<99) pressScore = 2;
        else if(sbp>=160 && dbp>=100) pressScore = 3;

        if(sick && sick.indexOf("糖尿病")>=0) sickScore = 2;
        if(smoke!="從未吸菸") smokeScore = 2;

      }else if(sex=="女"){
        if(age<=34) ageScore = -9;
        else if(age<=39) ageScore = -4;
        else if(age<=44) ageScore = 0;
        else if(age<=49) ageScore = 3;
        else if(age<=54) ageScore = 6;
        else if(age<=59) ageScore = 7;
        else if(age<=64) ageScore = 8;
        else if(age<=69) ageScore = 8;
        else if(age<=74) ageScore = 8;

        if(cholesterol<160) cholesterolScore = -2;
        else if(cholesterol<=199) cholesterolScore = 0;
        else if(cholesterol<=239) cholesterolScore = 1;
        else if(cholesterol<=279) cholesterolScore = 1;
        else if(cholesterol>=280) cholesterolScore = 3;

        if(highCholesterol<35) highCholesterolScore = 5;
        else if(highCholesterol<=44) highCholesterolScore = 2;
        else if(highCholesterol<=49) highCholesterolScore = 1;
        else if(highCholesterol<=59) highCholesterolScore = 0;
        else if(highCholesterol>=60) highCholesterolScore = -3;

        if(sbp<120 && dbp<80) pressScore = -3;
        else if(sbp<129 || dbp<84) pressScore = 0;
        else if(sbp<139 || dbp<89) pressScore = 0;
        else if(sbp<159 || dbp<99) pressScore = 2;
        else if(sbp>=160 && dbp>=100) pressScore = 3;

        if(sick && sick.indexOf("糖尿病")>=0) sickScore = 4;
        if(smoke!="從未吸菸") smokeScore = 2;
      }
      totaScore = ageScore + cholesterolScore + highCholesterolScore + pressScore + sickScore + smokeScore;
    
      if(sex=="男"){
        if(totaScore<-1) probability = "2%"
        else if(totaScore==0) probability = "3%"
        else if(totaScore==1) probability = "3%"
        else if(totaScore==2) probability = "4%"
        else if(totaScore==3) probability = "5%"
        else if(totaScore==4) probability = "7%"
        else if(totaScore==5) probability = "8%"
        else if(totaScore==6) probability = "10%"
        else if(totaScore==7) probability = "13%"
        else if(totaScore==8) probability = "16%"
        else if(totaScore==9) probability = "20%"
        else if(totaScore==10) probability = "25%"
        else if(totaScore==11) probability = "31%"
        else if(totaScore==12) probability = "37%"
        else if(totaScore==13) probability = "45%"
        else if(totaScore>=14) probability = "53%"
      }else if(sex=="女"){
        if(totaScore<=-2) probability = "1%"
        else if(totaScore==-1) probability = "2%"
        else if(totaScore==0) probability = "2%"
        else if(totaScore==1) probability = "2%"
        else if(totaScore==2) probability = "3%"
        else if(totaScore==3) probability = "3%"
        else if(totaScore==4) probability = "4%"
        else if(totaScore==5) probability = "4%"
        else if(totaScore==6) probability = "5%"
        else if(totaScore==7) probability = "6%"
        else if(totaScore==8) probability = "7%"
        else if(totaScore==9) probability = "8%"
        else if(totaScore==10) probability = "10%"
        else if(totaScore==11) probability = "11%"
        else if(totaScore==12) probability = "13%"
        else if(totaScore==13) probability = "15%"
        else if(totaScore==14) probability = "18%"
        else if(totaScore==15) probability = "20%"
        else if(totaScore==16) probability = "24%"
        else if(totaScore>=17) probability = ">=27%"
      }
    
    }else{
      totaScore = ""
    }
    
    

    arrayList.push({
      "身份證字號": user["身份證字號"],
      "姓名": user["中文姓名"],
      "性別": sex,
      "年齡": age,
      "年齡分數": ageScore,
      "膽固醇": cholesterol,
      "膽固醇分數": cholesterolScore,
      "高密度膽固醇": highCholesterol,
      "高密度膽固醇分數": highCholesterolScore,
      "收縮壓": sbp,
      "舒張壓": dbp,
      "血壓分數": pressScore,
      "糖尿病分數": sickScore,
      "抽菸分數": smokeScore,
      "總分數": totaScore,
      "十年內發生缺血性心臟病的機率":probability
    })

  })

  return arrayList

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

  const heartArrayList = await parserHeartExcel(filename)
  const heartWs = xlsx.utils.json_to_sheet(heartArrayList)
  xlsx.utils.book_append_sheet(excel,heartWs,"心力評量表")

  const id = crypto.randomBytes(20).toString('hex');
  xlsx.writeFile(excel,`result/${id}.xlsx`)
  return `${id}.xlsx`
}