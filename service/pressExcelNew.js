const xlsx = require('xlsx');
const XLSXStyle = require('xlsx-style');
const crypto = require('crypto');
const XlsxPopulate = require('xlsx-populate');
const fs = require('fs');
const pressExcel = require('./pressExcel');

async function fileExists(filePath) {
  return new Promise((resolve, reject) => {
    fs.exists(filePath, (exists) => {
      resolve(exists);
    });
  }).catch((error) => {
    console.log(error);
  });
}

async function getFile(filePath) {
  const exist = await fileExists(filePath);
  if (exist) {
    return readFile(filePath);
  }
  return '{}';
}

async function readFile(filePath) {
  return new Promise((resolve, reject) => {
    fs.readFile(filePath, (error, data) => {
      if (error) reject(error);
      resolve(data);
    });
  }).catch((error) => {
    console.log(error);
  });
}
function getPressExcelOutput(sunCount, totalPeople) {
  let result = {};
  let noCount = totalPeople - sunCount.countPeople;
  let fourNoCount = totalPeople - sunCount.four.countPeople;
  let unNormal = sunCount.first + sunCount.second + sunCount.third;
  let fourUnNormal = sunCount.four.first + sunCount.four.second + sunCount.four.third;
  result = {
    測量人數: sunCount.countPeople,
    未測量: noCount,
    測量率: Math.round((sunCount.countPeople / totalPeople) * 10000) / 100.0 + '%',
    未測量率: Math.round((noCount / totalPeople) * 10000) / 100.0 + '%',
    '測量次數<8': sunCount.lessThanPeople,
    總量測次數: sunCount.totalCount,
    平均測量次數: Math.round((sunCount.totalCount / sunCount.countPeople) * 10) / 10.0 + '次',
    正常: sunCount.normal,
    血壓異常率: Math.round((unNormal / sunCount.countPeople) * 10000) / 100.0 + '%',
    高血壓前期: sunCount.early,
    第一期: sunCount.first,
    第二期: sunCount.second,
    高血壓危象: sunCount.third,
    計畫測量人數: sunCount.four.countPeople,
    計畫未測量人數: fourNoCount,
    '計畫測量率%': Math.round((sunCount.four.countPeople / totalPeople) * 10000) / 100.0 + '%',
    '計畫未測量率%': Math.round((fourNoCount / totalPeople) * 10000) / 100.0 + '%',
    計畫總量測次數: sunCount.four.totalCount,
    門店平均量測次數: sunCount.four.countPeople == 0 ? '0次' : Math.round((sunCount.four.totalCount / sunCount.four.countPeople) * 10) / 10.0 + '次',
    計畫正常: sunCount.four.normal,
    計畫前期: sunCount.four.early,
    計畫第一期: sunCount.four.first,
    計畫第二期: sunCount.four.second,
    計畫危象: sunCount.four.third,
    計畫血壓異常率: sunCount.four.countPeople == 0 ? '0%' : Math.round((fourUnNormal / sunCount.four.countPeople) * 10000) / 100.0 + '%',
  };
  return result;
}
function getPressSumCount(dataList, column) {
  let tpmResult = {};
  for (let i of dataList) {
    if (i['量測次數'] && !i['未測量原因']) {
      let departObj = tpmResult[i[column]];
      if (!departObj) {
        departObj = {
          countPeople: 0,
          lessThanPeople: 0,
          totalCount: 0,
          normal: 0,
          early: 0,
          first: 0,
          second: 0,
          third: 0,
          four: {
            countPeople: 0,
            totalCount: 0,
            normal: 0,
            early: 0,
            first: 0,
            second: 0,
            third: 0,
          },
        };
      }
      departObj.countPeople = departObj.countPeople + 1;
      if (i['量測次數'] < 8) departObj.lessThanPeople = departObj.lessThanPeople + 1;
      departObj.totalCount = departObj.totalCount + i['量測次數'] * 1;
      if (i['血壓分級'] == '正常') departObj.normal = departObj.normal + 1;
      else if (i['血壓分級'] == '前期') departObj.early = departObj.early + 1;
      else if (i['血壓分級'] == '一期') departObj.first = departObj.first + 1;
      else if (i['血壓分級'] == '二期') departObj.second = departObj.second + 1;
      else if (i['血壓分級'] == '危象') departObj.third = departObj.third + 1;
      if (i['量測次數'] >= 4) {
        departObj.four.countPeople = departObj.four.countPeople + 1;
        departObj.four.totalCount = departObj.four.totalCount + i['量測次數'] * 1;
        if (i['血壓分級'] == '正常') departObj.four.normal = departObj.four.normal + 1;
        else if (i['血壓分級'] == '前期') departObj.four.early = departObj.four.early + 1;
        else if (i['血壓分級'] == '一期') departObj.four.first = departObj.four.first + 1;
        else if (i['血壓分級'] == '二期') departObj.four.second = departObj.four.second + 1;
        else if (i['血壓分級'] == '危象') departObj.four.third = departObj.four.third + 1;
      }
      tpmResult[i[column]] = departObj;
    }
  }
  return tpmResult;
}
function getTwoList(dataList, removeIdData) {
  let tpmResult = getPressSumCount(dataList, '第一層組織');

  let result = [];
  for (let i in tpmResult) {
    let totalPeople = 0;
    for (let j of dataList) {
      if (j['第一層組織'] == i) {
        const id = j['工號'];
        if (!removeIdData[id]) {
          totalPeople = totalPeople + 1;
        }
      }
    }

    if (totalPeople != 0) {
      const excelOutput = getPressExcelOutput(tpmResult[i], totalPeople);

      result.push({
        處: i,
        總人數: totalPeople,
        ...excelOutput,
      });
    }
  }
  return result;
}

function getOneList(dataList, removeIdData) {
  let tpmResult = getPressSumCount(dataList, '第二層組織');

  let result = [];
  for (let i in tpmResult) {
    let two = '';
    let totalPeople = 0;
    for (let j of dataList) {
      if (j['第二層組織'] == i) {
        const id = j['工號'];
        if (!removeIdData[id]) {
          totalPeople = totalPeople + 1;
          two = j['第一層組織'];
        }
      }
    }

    if (totalPeople != 0) {
      const excelOutput = getPressExcelOutput(tpmResult[i], totalPeople);

      result.push({
        處: two,
        區: i,
        總人數: totalPeople,
        ...excelOutput,
      });
    }
  }
  return result;
}

function getDepartList(dataList, removeIdData) {
  let tpmResult = getPressSumCount(dataList, '區域');

  let result = [];
  for (let i in tpmResult) {
    let two = '';
    let one = '';
    let departNo = '';
    let totalPeople = 0;

    for (let j of dataList) {
      if (j['區域'] == i) {
        const id = j['工號'];
        if (!removeIdData[id]) {
          totalPeople = totalPeople + 1;
          two = j['第一層組織'];
          one = j['第二層組織'];
          departNo = j['區域單位代碼'];
        }
      }
    }

    if (totalPeople != 0) {
      const excelOutput = getPressExcelOutput(tpmResult[i], totalPeople);

      result.push({
        處: two,
        區: one,
        代碼: departNo,
        部門名稱: i,
        總人數: totalPeople,
        ...excelOutput,
      });
    }
  }
  return result;
}

module.exports.parserPressExcel = async function parserPressExcel(filename) {
  const removeFile = await getFile(`removeId.json`);
  const removeIdData = JSON.parse(removeFile);

  const excel = xlsx.readFile(filename);
  var xlData = xlsx.utils.sheet_to_json(excel.Sheets['血壓統計']);
  let result = [];
  xlData.forEach((user) => {
    const workId = user['工號'];
    const reason = removeIdData[workId] ? removeIdData[workId].reason : '';
    const userUser = {
      ...user,
      未測量原因: reason,
    };
    result.push(userUser);
  });

  const ws = xlsx.utils.json_to_sheet(result);
  xlsx.utils.book_append_sheet(excel, ws, '血壓統計(有加上未測量原因)');
  changePressColor(ws, 15, result.length);

  // 處
  const twoData = getTwoList(result, removeIdData);
  const twoWs = xlsx.utils.json_to_sheet(twoData);
  xlsx.utils.book_append_sheet(excel, twoWs, '處');
  pressExcel.changePressTwoColor(twoWs, 27, twoData.length);

  // 區
  const oneData = getOneList(result, removeIdData);
  const oneWs = xlsx.utils.json_to_sheet(oneData);
  xlsx.utils.book_append_sheet(excel, oneWs, '區');
  pressExcel.changePressOneColor(oneWs, 28, oneData.length);

  // 店
  const departData = getDepartList(result, removeIdData);
  const departWs = xlsx.utils.json_to_sheet(departData);
  xlsx.utils.book_append_sheet(excel, departWs, '店');
  pressExcel.changePressDepartColor(departWs, 30, departData.length);

  const id = crypto.randomBytes(20).toString('hex');
  XLSXStyle.writeFile(excel, `result/${id}.xlsx`);

  return `${id}.xlsx`;
};

const ABCArr = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH'];

function changePressColor(ws, columnLength, dataLength) {
  let wscols = [
    { wch: 20 }, //A
    { wch: 20 }, //B
    { wch: 20 }, //C
    { wch: 11 }, //D
    { wch: 11 }, //E
    { wch: 11 }, //F
    { wch: 11 }, //G
    { wch: 11 }, //H
    { wch: 11 }, //I
    { wch: 11 }, //J
    { wch: 11 }, //K
    { wch: 11 }, //L
    { wch: 11 }, //M
    { wch: 11 }, //N
    { wch: 20 }, //O
  ];

  ws['!cols'] = wscols;

  // title
  for (let i = 0; i < columnLength; i++) {
    ws[`${ABCArr[i]}1`].s = {
      fill: {
        fgColor: { rgb: 'FFFFFACD' },
      },
      alignment: {
        horizontal: 'center',
      },
    };
  }

  // data
  for (let i = 0; i < columnLength; i++) {
    for (let j = 2; j <= dataLength + 1; j++) {
      if (ws[`${ABCArr[i]}${j}`]) {
        ws[`${ABCArr[i]}${j}`].s = {
          alignment: {
            horizontal: 'center',
          },
        };
      }
    }
  }
}
