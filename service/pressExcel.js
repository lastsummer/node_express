const xlsx = require('xlsx');
const crypto = require('crypto');
const XlsxPopulate = require('xlsx-populate');
const fs = require('fs')

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

function addSevenZero(str){
  if(str.length<2){
    return `000000${str}`
  }else if(str.length<3){
    return `00000${str}`
  }else if(str.length<4){
    return `0000${str}`
  }else if(str.length<5){
    return `000${str}`
  }else if(str.length<6){
    return `00${str}`
  }else if(str.length<7){
    return `0${str}`
  }
  return str
}

function reduceSevenZero(str){
  if(str.length>=8){
    if(str.substr(0, str.length-7)*1==0) return  str.substr(str.length-8, 7)
  }
  return str
}

function checkIDCorrect(str){
  let result = true;
  if(str.length>=8){
    result = false
  }
  return result
}

function getPressSumCount(dataList, column){
  let tpmResult = {}
  for (let i of dataList) {
    let departObj = tpmResult[i[column]]

    if(!departObj){
      departObj = {
        countPeople:0,
        lessThanPeople:0,
        totalCount: 0,
        normal: 0,
        early: 0,
        first: 0,
        second: 0,
        third: 0,
        four: {
          countPeople:0,
          totalCount: 0,
          normal: 0,
          early: 0,
          first: 0,
          second: 0,
          third: 0,
        }
      }
    }
    departObj.countPeople = departObj.countPeople + 1;
    if(i["測量次數"]<8) departObj.lessThanPeople = departObj.lessThanPeople + 1;
    departObj.totalCount = departObj.totalCount + (i["測量次數"])*1;
    if(i["高血壓等級"]=="正常") departObj.normal = departObj.normal + 1;
    else if(i["高血壓等級"]=="高血壓前期") departObj.early = departObj.early + 1;
    else if(i["高血壓等級"]=="第一期") departObj.first = departObj.first + 1;
    else if(i["高血壓等級"]=="第二期") departObj.second = departObj.second + 1;
    else if(i["高血壓等級"]=="高血壓危象") departObj.third = departObj.third + 1;
    if(i["測量次數"]>=4){
      departObj.four.countPeople = departObj.four.countPeople + 1;
      departObj.four.totalCount = departObj.four.totalCount + (i["測量次數"])*1;
      if(i["高血壓等級"]=="正常") departObj.four.normal = departObj.four.normal + 1;
      else if(i["高血壓等級"]=="高血壓前期") departObj.four.early = departObj.four.early + 1;
      else if(i["高血壓等級"]=="第一期") departObj.four.first = departObj.four.first + 1;
      else if(i["高血壓等級"]=="第二期") departObj.four.second = departObj.four.second + 1;
      else if(i["高血壓等級"]=="高血壓危象") departObj.four.third = departObj.four.third + 1;
    } 

    tpmResult[i[column]] = departObj
  }
  return tpmResult
}

function getPressExcelOutput(sunCount, totalPeople){
  let result = {}
  let noCount = totalPeople - sunCount.countPeople
  let fourNoCount = totalPeople - sunCount.four.countPeople
  let unNormal = sunCount.first + sunCount.second + sunCount.third
  let fourUnNormal = sunCount.four.first + sunCount.four.second + sunCount.four.third
  result = {
    測量人數: sunCount.countPeople,
    未測量: noCount,
    測量率: (Math.round(sunCount.countPeople / totalPeople * 10000) / 100.00) + "%",
    未測量率: (Math.round(noCount / totalPeople * 10000) / 100.00) + "%",
    "測量次數<8": sunCount.lessThanPeople,
    總量測次數: sunCount.totalCount,
    平均測量次數: (Math.round(sunCount.totalCount / sunCount.countPeople * 10) / 10.00) + "次" ,
    正常: sunCount.normal,
    血壓異常率: (Math.round(unNormal/sunCount.countPeople * 10000) / 100.00) + "%",
    高血壓前期: sunCount.early,
    第一期: sunCount.first,
    第二期: sunCount.second,
    高血壓危象: sunCount.third,
    計畫測量人數: sunCount.four.countPeople,
    計畫未測量人數: fourNoCount,
    "計畫測量率%": (Math.round(sunCount.four.countPeople / totalPeople * 10000) / 100.00) + "%",
    "計畫未測量率%": (Math.round(fourNoCount / totalPeople * 10000) / 100.00) + "%",
    計畫總量測次數: sunCount.four.totalCount,
    門店平均量測次數: (Math.round(sunCount.four.totalCount / sunCount.four.countPeople * 10) / 10.00) + "次" ,
    計畫正常: sunCount.four.normal,
    計畫前期: sunCount.four.early,
    計畫第一期: sunCount.four.first,
    計畫第二期: sunCount.four.second,
    計畫危象: sunCount.four.third,
    計畫血壓異常率: (Math.round(fourUnNormal/sunCount.four.countPeople * 10000) / 100.00) + "%",
  }
  return result
}

function getDepartList(dataList, idData){
  let tpmResult = getPressSumCount(dataList, "部門名稱")

  let result = []
  for (let i in tpmResult) {
    let two = ""
    let one = ""
    let departNo = ""
    let totalPeople = 0
    for (let id in idData) {
      if(idData[id].depart == i){
        two = idData[id].two
        one = idData[id].one
        departNo = idData[id].departNo
        totalPeople = totalPeople + 1
      } 
    }

    if(totalPeople!=0){
      const excelOutput = getPressExcelOutput(tpmResult[i], totalPeople)
  
      result.push({
        處: two,
        區: one,
        代碼: departNo,
        部門名稱: i,
        總人數: totalPeople,
        ...excelOutput
      })
    }
  }
  return result
}

function getOneList(dataList, idData){
  let tpmResult = getPressSumCount(dataList, "上一層組織中文名稱")

  let result = []
  for (let i in tpmResult) {
    let two = ""
    let totalPeople = 0
    for (let id in idData) {
      if(idData[id].one == i){
        two = idData[id].two
        totalPeople = totalPeople + 1
      } 
    }

    if(totalPeople!=0){
      const excelOutput = getPressExcelOutput(tpmResult[i], totalPeople)
  
      result.push({
        處: two,
        區: i,
        總人數: totalPeople,
        ...excelOutput
      })
    }
  }
  return result
}

function getTwoList(dataList, idData){
  let tpmResult = getPressSumCount(dataList, "上兩層組織中文名稱")

  let result = []
  for (let i in tpmResult) {
    let totalPeople = 0
    for (let id in idData) {
      if(idData[id].two == i){
        totalPeople = totalPeople + 1
      } 
    }

    if(totalPeople!=0){
      const excelOutput = getPressExcelOutput(tpmResult[i], totalPeople)
  
      result.push({
        處: i,
        總人數: totalPeople,
        ...excelOutput
      })
    }
  }
  return result
}

module.exports.parserPressExcel = async function parserPressExcel(filename, month) {
  let monthStr = ""
  if(month) monthStr = `-${month}`
  const file = await getFile(`workIdTOdepart${monthStr}.json`)
  const idData = JSON.parse(file);

  const excel = xlsx.readFile(filename);
  var xlData = xlsx.utils.sheet_to_json(excel.Sheets['表單回應 1']);
  let arrayList = {};
  xlData.forEach((user) => {

    const originWorkId = user['員工工號-7碼(若不足7碼，請前面補打【0】數字)']
    let name = user['姓名(全名)']
    let departName = user['請點選所在門市'] || user['請點選所在門市_2'] || user['請點選所在門市_3']
                   || user['請點選所在門市_4'] || user['請點選所在門市_5'] || user['請點選所在門市_6']
                   || user['請點選所在門市_7'] || user['請點選所在門市_8'] || user['請點選所在門市_9']
                   || user['請點選所在門市_10'] || user['請點選所在門市_11'] || user['請點選所在門市_12']
                   || user['請點選所在門市_13'] || user['請點選所在門市_1']
    let workId = addSevenZero(originWorkId+"")
    workId = reduceSevenZero(workId)
    let memo = ""
    // const isIDCorrect = checkIDCorrect(workId)
    let nameCount = 0
    let departCount = 0
    if(!idData[workId]){
      let nameWorkId = ""
      for (let id in idData) {
        if(idData[id].name == name){
          nameCount = nameCount+1
          nameWorkId = id
          memo = memo + "、" + id
        } 
      }
      if(nameCount>=2){
        let departWorkId = ""
        if(departName){
          for (let id in idData) {
            if(idData[id].name == name && idData[id].depart.indexOf(departName.substr(2, departName.length-2))>0 ){
              departCount = departCount+1
              departWorkId = id
            } 
          }
        }
        if(departCount==1){
          workId = departWorkId
        }else{
          workId = originWorkId
        }
      }else{
        workId = nameWorkId
      } 
    }else{
      name = idData[workId].name
    }
    
    const sbp = (user['收縮壓'])*1
    const dbp = (user['舒張壓'])*1
    const pulse = (user['脈搏'])*1
    const workName = user['職稱']
    const workArea = user['工作區域 ']
    if(sbp && sbp<=280 && sbp>=60
      && dbp && dbp<=200 && dbp>=30
      && pulse){
      const count = arrayList[workId] ? (arrayList[workId].count)*1 + 1 : 1
      const averageSbp = arrayList[workId] ? ((arrayList[workId].averageSbp)*1 + sbp) : sbp
      const averageDbp = arrayList[workId] ? ((arrayList[workId].averageDbp)*1 + dbp) : dbp
      const averagePulse = arrayList[workId] ? ((arrayList[workId].averagePulse)*1 + pulse) : pulse
      arrayList[workId] = { count, averageSbp, averageDbp, 
        averagePulse, nameCount, memo, 
        name, departCount, workName, workArea }
    }
  })

  let result = []
  for (let i in arrayList) {
    let averageSbp = Math.round(arrayList[i].averageSbp/ arrayList[i].count)
    if(averageSbp>=180) arrayList[i].desc = "高血壓危象"
    else if(averageSbp>=160) arrayList[i].desc = "第二期"
    else if(averageSbp>=140) arrayList[i].desc = "第一期"
    else if(averageSbp>=120) arrayList[i].desc = "高血壓前期"
    else arrayList[i].desc = "正常"
    result.push({
      上兩層組織中文名稱: idData[i]? idData[i].two: '',
      上一層組織中文名稱: idData[i]? idData[i].one: '',
      代碼: idData[i]? idData[i].departNo: '',
      部門名稱: idData[i]? idData[i].depart: '',
      工號: i,
      姓名: arrayList[i].name,
      職稱: arrayList[i].workName,
      工作區域: arrayList[i].workArea,
      平均收縮壓: Math.round(arrayList[i].averageSbp/ arrayList[i].count),
      平均舒張壓: Math.round(arrayList[i].averageDbp/ arrayList[i].count),
      平均脈搏: Math.round(arrayList[i].averagePulse/ arrayList[i].count),
      測量次數: arrayList[i].count,
      高血壓等級: arrayList[i].desc,
      備註: (arrayList[i].nameCount>=2 && arrayList[i].departCount!=1) ? arrayList[i].memo : ''
    })
  }

  // 血壓統計 需要加上沒有測量的人
  let existDepart = {}
  for (let i of result) {
    existDepart[i["代碼"]] = i["部門名稱"]
  }
  let pressResult = [...result]
  for (let id in idData) {
    if(existDepart[idData[id].departNo]){
      let isExist = false
      for (let i of result) {
        if(i["工號"]==id) isExist = true
      }
      if(!isExist){
        pressResult.push(
          {
            上兩層組織中文名稱: idData[id].two,
            上一層組織中文名稱: idData[id].one,
            代碼: idData[id].departNo,
            部門名稱: idData[id].depart,
            工號: id+"",
            姓名: idData[id].name
          }
        )
      }
    }
  }


  const ws = xlsx.utils.json_to_sheet(pressResult);
  // var workbook = xlsx.utils.book_new()
  xlsx.utils.book_append_sheet(excel, ws, '血壓統計');

  // 處
  const twoWs = xlsx.utils.json_to_sheet(getTwoList(result, idData))
  xlsx.utils.book_append_sheet(excel, twoWs, '處');

  // 區
  const oneWs = xlsx.utils.json_to_sheet(getOneList(result, idData))
  xlsx.utils.book_append_sheet(excel, oneWs, '區');

  // 店
  const departWs = xlsx.utils.json_to_sheet(getDepartList(result, idData))
  xlsx.utils.book_append_sheet(excel, departWs, '店');

  const id = crypto.randomBytes(20).toString('hex');
  xlsx.writeFile(excel, `result/${id}.xlsx`);

  await changePressColor(id)
  return `${id}.xlsx`;
}

async function changePressColor(fileName){
  return new Promise((resolve, reject) => {
    XlsxPopulate.fromFileAsync(`result/${fileName}.xlsx`)
    .then((workbook) => {
      const sheet = workbook.sheet('血壓統計');
      sheet.column("A").width(20)
      sheet.column("B").width(20)
      sheet.column("C").width(15)
      sheet.column("D").width(20)
      sheet.column("E").width(11)
      sheet.column("F").width(11)
      sheet.column("L").width(30)
      const rows = sheet._rows;
      rows.forEach((row) => {
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

      const twoSheet = workbook.sheet('處');
      twoSheet.column("A").width(20)
      twoSheet.column("H").width(15)
      const towRows = twoSheet._rows;
      towRows.forEach((row) => {
        row._cells.forEach((cell) => {
          let style = {
            horizontalAlignment: 'center'
          }
          if(cell.rowNumber()==1){
            if(cell.columnNumber()>=16){
              style.fill = '77DDFF'
            }else{
              style.fill = 'fffacd'
            }
          }
          cell.style(style)
        });
      });

      const oneSheet = workbook.sheet('區');
      oneSheet.column("A").width(20)
      oneSheet.column("B").width(20)
      oneSheet.column("I").width(15)
      const oneRows = oneSheet._rows;
      oneRows.forEach((row) => {
        row._cells.forEach((cell) => {
          let style = {
            horizontalAlignment: 'center'
          }
          if(cell.rowNumber()==1){
            if(cell.columnNumber()>=17){
              style.fill = '77DDFF'
            }else{
              style.fill = 'fffacd'
            }
          }
          cell.style(style)
        });
      });

      const departSheet = workbook.sheet('店');
      departSheet.column("A").width(20)
      departSheet.column("B").width(20)
      departSheet.column("C").width(15)
      departSheet.column("D").width(25)
      departSheet.column("K").width(15)
      const departRows = departSheet._rows;
      departRows.forEach((row) => {
        row._cells.forEach((cell) => {
          let style = {
            horizontalAlignment: 'center'
          }
          if(cell.rowNumber()==1){
            if(cell.columnNumber()>=19){
              style.fill = '77DDFF'
            }else{
              style.fill = 'fffacd'
            }
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