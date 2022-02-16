const xlsx = require('xlsx');
const crypto = require('crypto');
const XlsxPopulate = require('xlsx-populate');
const fs = require('fs')

function formatDate(numb) {
  if (numb != undefined) {
    let time = new Date((numb - 1) * 24 * 3600000 + 1);
    time.setYear(time.getFullYear() - 70);
    return time.getFullYear();
  }
}


async function fileExists(filePath){
  return new Promise((resolve, reject) => {
    fs.exists(filePath, (exists) => { 
      resolve(exists);
    });
  }).catch(error => {
    console.log(error);
  })
}

async function getFile(filePath){
  const exist = await fileExists(filePath);
  if(exist){
    return readFile(filePath);
  }
  return "{}";
}

async function readFile(filePath){
  return new Promise((resolve, reject) => {
    fs.readFile(filePath, (error, data) => {
      if(error)reject(error)
      resolve(data)
    })
  }).catch(error => {
    console.log(error);
  })
}

function validateIdNumberToAgeYear(str) {
  let date = new Date();
  let currentYear = date.getFullYear();
  let birdthdayArr = str.split('-');
  let year = birdthdayArr[0];
  if (birdthdayArr.length == 1) {
    birdthdayArr = str.split('/');
    year = birdthdayArr[0];
  }
  if (birdthdayArr.length == 1) { // excel 預設日期格式
    year = formatDate(str);
  }
  if(year>currentYear){
    year = (str.substr(0, str.length-4))*1 + 1911 
  }
  return currentYear - year;
}
async function parserHeartExcel(filename) {

  const file = await getFile(`idTOdepart.json`)
  const idData = JSON.parse(file);
  
  const excel = xlsx.readFile(filename);
  var xlData = xlsx.utils.sheet_to_json(excel.Sheets['全聯實業股份有限公司總表資料']);
  let arrayList = [];
  xlData.forEach((user) => {
    let totaScore = 0;
    let ageScore = 0;
    let cholesterolScore = 0;
    let highCholesterolScore = 0;
    let pressScore = 0;
    let sickScore = 0;
    let smokeScore = 0;
    let probability = '';
    const age = validateIdNumberToAgeYear(user['出生日期'] + '');
    const sex = user['性別'];
    const cholesterol = user['總膽固醇'];
    const highCholesterol = user['高密度-脂蛋白'];
    const sbp = user['收縮壓'];
    const dbp = user['舒張壓'];
    const sick = user['慢性病史'];
    const smoke = user['抽菸'];
    if (age != '' && age >= 30) {
      if (sex == '男') {
        if (age <= 34) ageScore = -1;
        else if (age <= 39) ageScore = 0;
        else if (age <= 44) ageScore = 1;
        else if (age <= 49) ageScore = 2;
        else if (age <= 54) ageScore = 3;
        else if (age <= 59) ageScore = 4;
        else if (age <= 64) ageScore = 5;
        else if (age <= 69) ageScore = 6;
        else if (age <= 74) ageScore = 7;

        if (cholesterol < 160) cholesterolScore = -3;
        else if (cholesterol <= 199) cholesterolScore = 0;
        else if (cholesterol <= 239) cholesterolScore = 1;
        else if (cholesterol <= 279) cholesterolScore = 2;
        else if (cholesterol >= 280) cholesterolScore = 3;

        if (highCholesterol < 35) highCholesterolScore = 2;
        else if (highCholesterol <= 44) highCholesterolScore = 1;
        else if (highCholesterol <= 49) highCholesterolScore = 0;
        else if (highCholesterol <= 59) highCholesterolScore = 0;
        else if (highCholesterol >= 60) highCholesterolScore = -2;

        if (sbp < 120 && dbp < 80) pressScore = 0;
        else if (sbp < 129 || dbp < 84) pressScore = 0;
        else if (sbp < 139 || dbp < 89) pressScore = 1;
        else if (sbp < 159 || dbp < 99) pressScore = 2;
        else if (sbp >= 160 && dbp >= 100) pressScore = 3;

        if (sick && sick.indexOf('糖尿病') >= 0) sickScore = 2;
        if (smoke != '從未吸菸') smokeScore = 2;
      } else if (sex == '女') {
        if (age <= 34) ageScore = -9;
        else if (age <= 39) ageScore = -4;
        else if (age <= 44) ageScore = 0;
        else if (age <= 49) ageScore = 3;
        else if (age <= 54) ageScore = 6;
        else if (age <= 59) ageScore = 7;
        else if (age <= 64) ageScore = 8;
        else if (age <= 69) ageScore = 8;
        else if (age <= 74) ageScore = 8;

        if (cholesterol < 160) cholesterolScore = -2;
        else if (cholesterol <= 199) cholesterolScore = 0;
        else if (cholesterol <= 239) cholesterolScore = 1;
        else if (cholesterol <= 279) cholesterolScore = 1;
        else if (cholesterol >= 280) cholesterolScore = 3;

        if (highCholesterol < 35) highCholesterolScore = 5;
        else if (highCholesterol <= 44) highCholesterolScore = 2;
        else if (highCholesterol <= 49) highCholesterolScore = 1;
        else if (highCholesterol <= 59) highCholesterolScore = 0;
        else if (highCholesterol >= 60) highCholesterolScore = -3;

        if (sbp < 120 && dbp < 80) pressScore = -3;
        else if (sbp < 129 || dbp < 84) pressScore = 0;
        else if (sbp < 139 || dbp < 89) pressScore = 0;
        else if (sbp < 159 || dbp < 99) pressScore = 2;
        else if (sbp >= 160 && dbp >= 100) pressScore = 3;

        if (sick && sick.indexOf('糖尿病') >= 0) sickScore = 4;
        if (smoke != '從未吸菸') smokeScore = 2;
      }
      totaScore = ageScore + cholesterolScore + highCholesterolScore + pressScore + sickScore + smokeScore;

      if (sex == '男') {
        if (totaScore <= -1) probability = '2%';
        else if (totaScore == 0) probability = '3%';
        else if (totaScore == 1) probability = '3%';
        else if (totaScore == 2) probability = '4%';
        else if (totaScore == 3) probability = '5%';
        else if (totaScore == 4) probability = '7%';
        else if (totaScore == 5) probability = '8%';
        else if (totaScore == 6) probability = '10%';
        else if (totaScore == 7) probability = '13%';
        else if (totaScore == 8) probability = '16%';
        else if (totaScore == 9) probability = '20%';
        else if (totaScore == 10) probability = '25%';
        else if (totaScore == 11) probability = '31%';
        else if (totaScore == 12) probability = '37%';
        else if (totaScore == 13) probability = '45%';
        else if (totaScore >= 14) probability = '53%';
      } else if (sex == '女') {
        if (totaScore <= -2) probability = '1%';
        else if (totaScore == -1) probability = '2%';
        else if (totaScore == 0) probability = '2%';
        else if (totaScore == 1) probability = '2%';
        else if (totaScore == 2) probability = '3%';
        else if (totaScore == 3) probability = '3%';
        else if (totaScore == 4) probability = '4%';
        else if (totaScore == 5) probability = '4%';
        else if (totaScore == 6) probability = '5%';
        else if (totaScore == 7) probability = '6%';
        else if (totaScore == 8) probability = '7%';
        else if (totaScore == 9) probability = '8%';
        else if (totaScore == 10) probability = '10%';
        else if (totaScore == 11) probability = '11%';
        else if (totaScore == 12) probability = '13%';
        else if (totaScore == 13) probability = '15%';
        else if (totaScore == 14) probability = '18%';
        else if (totaScore == 15) probability = '20%';
        else if (totaScore == 16) probability = '24%';
        else if (totaScore >= 17) probability = '>=27%';
      }
    } else {
      totaScore = '';
    }

    arrayList.push({
      身份證字號: user['身份證字號'],
      姓名: user['中文姓名'],
      上兩層組織中文名稱: idData[user['身份證字號']]? idData[user['身份證字號']].two: '',
      上一層組織中文名稱: idData[user['身份證字號']]? idData[user['身份證字號']].one: '',
      部門名稱: idData[user['身份證字號']]? idData[user['身份證字號']].depart: '',
      性別: sex,
      年齡: age,
      年齡分數: ageScore,
      膽固醇: cholesterol,
      膽固醇分數: cholesterolScore,
      高密度膽固醇: highCholesterol,
      高密度膽固醇分數: highCholesterolScore,
      收縮壓: sbp,
      舒張壓: dbp,
      血壓分數: pressScore,
      糖尿病分數: sickScore,
      抽菸分數: smokeScore,
      總分數: totaScore,
      十年內發生缺血性心臟病的機率: probability,
    });
  });

  return arrayList;
}

module.exports.parserExcel = async function parserExcel(filename) {

  const file = await getFile(`idTOdepart.json`)
  const idData = JSON.parse(file);

  const excel = xlsx.readFile(filename);
  var xlData = xlsx.utils.sheet_to_json(excel.Sheets['全聯實業股份有限公司總表資料']);
  let arrayList = [];
  let weightArrayList = [];

  xlData.forEach((user) => {
    let incompatible = {
      body: '',
      sbp: '',
      dbp: '',
      waistline: '',
      proteinuria: '',
      leukocyte: '',
      hemoglobin: '',
      alt: '',
      creatinine: '',
      cholesterol: '',
      triglycerides: '',
      HDLC: '',
      LDLC: '',
      glucose: ''
    };
    let count = 0
    // 身體質量指數 9
    if (user['身體質量指數'] >= 35) {
      count = count+1
      incompatible.body = user['身體質量指數']
    }
    // 收縮壓 10
    if (user['收縮壓'] >= 160) {
      count = count+1
      incompatible.sbp = user['收縮壓']
    }
    // 舒張壓 11
    if (user['舒張壓'] >= 100) {
      count = count+1
      incompatible.dbp = user['舒張壓']
    }
    // 腰圍 12
    if (user['腰圍'] >= 90) {
      count = count+1
      incompatible.waistline = user['腰圍']
    }
    // 尿蛋白 13
    if (user['尿蛋白'] == '4+') {
      count = count+1
      incompatible.proteinuria = user['尿蛋白']
    }
    // 白血球 15
    if (user['白血球'] * 1000 >= 20000 || user['白血球'] * 1000 <= 2500) {
      count = count+1
      incompatible.leukocyte = user['白血球']
    }
    // 血色素 16
    if (user['血色素'] >= 21 || user['血色素'] <= 7) {
      count = count+1
      incompatible.hemoglobin = user['血色素']
    }
    // 丙氨酸轉氨脢 ALT 17
    if (user['丙氨酸轉氨脢 ALT'] >= 151) {
      count = count+1
      incompatible.alt = user['丙氨酸轉氨脢 ALT']
    }
    // 肌酸酐 18
    if (user['肌酸酐'] >= 2.5) {
      count = count+1
      incompatible.creatinine = user['肌酸酐']
    }
    // 總膽固醇 19
    if (user['總膽固醇'] >= 301) {
      count = count+1
      incompatible.cholesterol = user['總膽固醇']
    }
    // 三酸甘油脂 20
    if (user['三酸甘油脂'] >= 501) {
      count = count+1
      incompatible.triglycerides = user['三酸甘油脂']
    }
    // 高密度-脂蛋白 21
    if (user['高密度-脂蛋白'] <= 40) {
      count = count+1
      incompatible.HDLC = user['高密度-脂蛋白']
    }
    // 低密度-脂蛋白 22
    if (user['低密度-脂蛋白'] >= 191) {
      count = count+1
      incompatible.LDLC = user['低密度-脂蛋白']
    }
    // 空腹血糖 23
    if (user['空腹血糖'] >= 161) {
      count = count+1
      incompatible.glucose = user['空腹血糖']
    }

    let weight = ""
    if(user['身體質量指數']>=35){
      weight = "重度肥胖"
    }else if(user['身體質量指數']>=30){
      weight = "中度肥胖"
    }else if(user['身體質量指數']>=27){
      weight = "輕度肥胖"
    }else if(user['身體質量指數']>=24){
      weight = "過重"
    }else if(user['身體質量指數']>=18.5){
      weight = "正常"
    }else{
      weight = "體重過輕"
    }

    weightArrayList.push({
      身份證字號: user['身份證字號'],
      姓名: user['中文姓名'],
      上兩層組織中文名稱: idData[user['身份證字號']]? idData[user['身份證字號']].two: '',
      上一層組織中文名稱: idData[user['身份證字號']]? idData[user['身份證字號']].one: '',
      部門名稱: idData[user['身份證字號']]? idData[user['身份證字號']].depart: '',
      身體質量指數: user['身體質量指數'],
      重量: weight
    })

    arrayList.push({
      身份證字號: user['身份證字號'],
      姓名: user['中文姓名'],
      上兩層組織中文名稱: idData[user['身份證字號']]? idData[user['身份證字號']].two: '',
      上一層組織中文名稱: idData[user['身份證字號']]? idData[user['身份證字號']].one: '',
      部門名稱: idData[user['身份證字號']]? idData[user['身份證字號']].depart: '',
      身體質量指數: incompatible.body,
      收縮壓: incompatible.sbp,
      舒張壓: incompatible.dbp,
      腰圍: incompatible.waistline,
      尿蛋白: incompatible.proteinuria,
      白血球: incompatible.leukocyte,
      血色素: incompatible.hemoglobin,
      "丙氨酸轉氨脢 ALT": incompatible.alt,
      肌酸酐: incompatible.creatinine,
      總膽固醇: incompatible.cholesterol,
      三酸甘油脂: incompatible.triglycerides,
      "高密度-脂蛋白": incompatible.HDLC,
      "低密度-脂蛋白": incompatible.LDLC,
      空腹血糖: incompatible.glucose,
      符合項目: count,
    });
  });
  const ws = xlsx.utils.json_to_sheet(arrayList);
  xlsx.utils.book_append_sheet(excel, ws, '4級列表');

  const heartArrayList = await parserHeartExcel(filename);
  const heartWs = xlsx.utils.json_to_sheet(heartArrayList);
  xlsx.utils.book_append_sheet(excel, heartWs, '心力評量表');

  const weightWs = xlsx.utils.json_to_sheet(weightArrayList);
  xlsx.utils.book_append_sheet(excel, weightWs, '重量');

  const id = crypto.randomBytes(20).toString('hex');
  xlsx.writeFile(excel, `result/${id}.xlsx`);

  await changeColor(id)
  return `${id}.xlsx`;
};


async function changeColor(fileName){
  return new Promise((resolve, reject) => {
    XlsxPopulate.fromFileAsync(`result/${fileName}.xlsx`)
    .then((workbook) => {
      const sheet = workbook.sheet('4級列表');
      sheet.column("A").width(15)
      sheet.column("B").width(11)
      sheet.column("C").width(20)
      sheet.column("D").width(20)
      sheet.column("E").width(20)
      sheet.column("F").width(15)
      sheet.column("M").width(15)
      sheet.column("Q").width(15)
      sheet.column("R").width(15)
      const rows = sheet._rows;
      rows.forEach((row) => {
        row._cells.forEach((cell) => {
          let style = {
            horizontalAlignment: 'center'
          }
          if(cell.columnNumber()>=6 && cell.columnNumber()<=19 && cell.rowNumber()>=2 && cell.value()){
            style.fill = 'ffff00'
          }else if(cell.rowNumber()==1){
            style.fill = 'fffacd'
          }

          // 255 250 205
          cell.style(style)
        });
      });

      // 心力量表
      const sheetHeart = workbook.sheet('心力評量表');
      sheetHeart.column("A").width(15)
      sheetHeart.column("B").width(11)
      sheetHeart.column("C").width(20)
      sheetHeart.column("D").width(20)
      sheetHeart.column("E").width(20)
      sheetHeart.column("S").width(30)
      const rowsHeart = sheetHeart._rows;
      rowsHeart.forEach((row) => {
        row._cells.forEach((cell) => {
          let style = {
            horizontalAlignment: 'center'
          }
          if(cell.rowNumber()==1){
            style.fill = 'fffacd'
          }
          cell.style(style)
        });
      });

      // 重量
      const sheetWeight = workbook.sheet('重量');
      sheetWeight.column("A").width(15)
      sheetWeight.column("B").width(11)
      sheetWeight.column("C").width(20)
      sheetWeight.column("D").width(20)
      sheetWeight.column("E").width(20)
      const rowsWeight = sheetWeight._rows;
      rowsWeight.forEach((row) => {
        row._cells.forEach((cell) => {
          let style = {
            horizontalAlignment: 'center'
          }
          if(cell.rowNumber()==1){
            style.fill = 'fffacd'
          }else if(cell.columnNumber()>=7 && cell.value()!="正常"){
            style.fill = 'ffff00'
          }
          cell.style(style)
        });
      });

      workbook.toFileAsync(`result/${fileName}.xlsx`);
      resolve(fileName)

    })
    .catch((error) => {
      console.log(error);
    });
    
  }).catch(error => {
    console.log(error);
  })
}





