const fs = require('fs')


function addZero(number){
  if(number<10){
    return `0${number}`
  }
  return number
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

async function writeFile(filePath, data){
  return new Promise((resolve, reject) => {
    fs.writeFile(filePath, data, function (err) {
      if (err)
        reject(err);
      else
        resolve('Write operation complete.');
    })
  }).catch(error => {
    console.log(error);
  })
}

async function getFile(filePath){
  const exist = await fileExists(filePath);
  if(exist){
    return readFile(filePath);
  }else{
    await writeFile(filePath, "{}");
  }
  return "{}";
}

function getTime(start, end){
  let time = 0;
  let timeFormate = "";
  let timeColor = "";
  let timeBar = 0;

  if(start!="" && end!=""){
    let startArr = start.split(":");
    let endArr = end.split(":");
    let startTime = startArr[0]*60+startArr[1]*1
    let endTime = endArr[0]*60+endArr[1]*1

    time = endTime - startTime;
    let hour = parseInt(time/60)
    timeFormate = `${hour}小時${time-hour*60}分`
    if(startTime>endTime){
      timeColor = "bg-danger";
      timeBar = "25%";
    }else if(hour<8){
      timeColor = "bg-info";
      timeBar = `${parseInt(hour/10*100)}%`;
    }else if(hour<=9){
      timeColor = "bg-success";
      timeBar = `${parseInt(hour/10*100)}%`;
    }else{
      timeColor = "bg-warning";
      timeBar = `100%`;
    } 
  }
  

  return {
    time,
    timeFormate,
    timeColor,
    timeBar
  }
}

const weekArray = {
  1: "一",
  2: "二",
  3: "三",
  4: "四",
  5: "五",
  6: "六",
  0: "日"
}

async function getMonthData(year, month, userName){
  let dayList = []
  const currentMonth = month+1;
  let userFileName = ""
  if (userName) {
    userFileName = `-${userName}`
  }
  const file = await getFile(`wang/data/${year}-${addZero(currentMonth)}${userFileName}.json`)
  const dayData = JSON.parse(file);
  let totalTime = 0;
  for(let i = 1; i<=31; i++){
    const day = new Date(`${currentMonth}/${i}/${year}`)
    if(day.getMonth()==month){
      let start = dayData[i]? dayData[i].start : ""
      let end = dayData[i]? dayData[i].end : ""
      let time = getTime(start, end)
      totalTime = totalTime + time.time
      let dayObj = {
        formate: `${year}-${addZero(currentMonth)}-${addZero(day.getDate())}(${weekArray[day.getDay()]})`,
        day: i,
        start,
        end,
        time
      }
      dayList.push(dayObj)
    }
    let hour = parseInt(totalTime/60)
    totalTimeFormate = `${hour}小時${totalTime-hour*60}分`
  }
  return {currentMonth: addZero(currentMonth), currentYear: year, totalTimeFormate , dayList}
}
module.exports.getMonthData = getMonthData
module.exports.getCurrentMonth = async function getCurrentMonth(userName){
  const today = new Date()
  return getMonthData(today.getFullYear(), today.getMonth(), userName)
}

async function getExistTime(year, month, day){
  const fileName = `wang/data/${year}-${addZero(month)}.json`
  const file = await getFile(fileName)
  const dayData = JSON.parse(file)
  let newMonthData = {}
  let totalTime = 0;
  for (const [dayKey, dayValue] of Object.entries(dayData)) {
    if(dayKey!=day){
      newMonthData[dayKey] = dayValue
      let time = getTime(dayValue.start, dayValue.end)
      totalTime = totalTime + time.time
    }
  }
  return { newMonthData, totalTime, fileName}
}

module.exports.saveTime = async function saveTime(year, month, day, start, end){
  let { newMonthData, totalTime, fileName} = await getExistTime(year, month, day)

  let formateStartTime = addTimeZero(start);
  let formateEndTime = addTimeZero(end);
  let formateTime = {start:formateStartTime, end: formateEndTime}
  newMonthData[day] = formateTime
  let time = getTime(formateStartTime, formateEndTime)
  totalTime = totalTime + time.time

  let hour = parseInt(totalTime/60)
  totalTimeFormate = `${hour}小時${totalTime-hour*60}分`

  await writeFile(fileName, JSON.stringify(newMonthData))

  return {time, totalTimeFormate, ...formateTime}
}

function addTimeZero(time){
  let timeArr = time.split(":");
  return `${addZero(timeArr[0]*1)}:${addZero(timeArr[1]*1)}`
}

module.exports.deleteTime = async function deleteTime(year, month, day){
  let { newMonthData, totalTime, fileName} = await getExistTime(year, month, day)

  let hour = parseInt(totalTime/60)
  totalTimeFormate = `${hour}小時${totalTime-hour*60}分`

  await writeFile(fileName, JSON.stringify(newMonthData))
  return {totalTimeFormate}
}