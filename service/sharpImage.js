const sharp = require('sharp');
const crypto = require('crypto');

function fillMinutes(minutes){
  if(minutes<10) return `0${minutes}`
  return `${minutes}`
}
const picPath = './service/pic/'
module.exports.getMetadata = async function getMetadata(code) {
  // 場所代碼
  const codeSplit = code.split("");
  let compositeArr = []
  let totalLeft = 396;
  for (let i in codeSplit) {
    let picture = `${picPath}number${codeSplit[i]}.jpg`
    if(codeSplit[i]==' '){
      picture = `${picPath}numberSpace.jpg`
      totalLeft = totalLeft + 9;
    }else{
      totalLeft = totalLeft + 20;
      compositeArr.push({ input: picture, top: 1498, left: totalLeft });
    } 
  }

  // 日期
  const timeTop = 1437
  const currentDate = new Date();
  const currentHour = currentDate.getHours() + 8;
  let picture = "morning.jpg"
  let hourStr = `${currentHour}`
  if(currentHour > 12){
    hourStr = `${currentHour - 12}`
    picture = 'afternoon.jpg'
  }else if(currentHour == 12){
    picture = 'afternoon.jpg'
  }
  const hourSplit = hourStr.split("");
  let timeTotalLeft = 478;
  for (let i in hourSplit) {
    let picture = `${picPath}time${hourSplit[i]}.jpg`
    timeTotalLeft = timeTotalLeft + i*15
    compositeArr.push({ input: picture, top: timeTop, left: timeTotalLeft });
  }
  timeTotalLeft = timeTotalLeft + 15
  compositeArr.push({ input: `${picPath}timeColon.jpg`, top: timeTop, left: timeTotalLeft });

  timeTotalLeft = timeTotalLeft + 9
  let minuteStr = `${fillMinutes(currentDate.getMinutes())}`
  const minuteSplit = minuteStr.split("");
  for (let i in minuteSplit) {
    let picture = `${picPath}time${minuteSplit[i]}.jpg`
    timeTotalLeft = timeTotalLeft + i*15
    compositeArr.push({ input: picture, top: timeTop, left: timeTotalLeft });
  }


  const fileName = crypto.randomBytes(20).toString('hex');

  const metadata = await sharp(`${picture}`)
    .composite(compositeArr)
    .png()
    .toFile(`./result/${fileName}.jpg`);

  return `${fileName}.jpg`;
};
