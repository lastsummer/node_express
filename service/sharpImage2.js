const sharp = require('sharp');
const crypto = require('crypto');

function fillMinutes(minutes){
  if(minutes<10) return `0${minutes}`
  return `${minutes}`
}
async function getMetadata(code) {
  // 場所代碼
  const codeSplit = code.split("");
  let compositeArr = []
  let totalLeft = 396;
  for (let i in codeSplit) {
    let picture = `./pic/number${codeSplit[i]}.jpg`
    if(codeSplit[i]==' '){
      picture = `./pic/numberSpace.jpg`
      totalLeft = totalLeft + 7;
    }else{
      totalLeft = totalLeft + 21;
      compositeArr.push({ input: picture, top: 1498, left: totalLeft });
    } 
  }

  // 日期
  const timeTop = 1437
  const currentDate = new Date();
  let timeText = `今天 上午 ${currentDate.getHours()}:${currentDate.getMinutes()}`;
  if (currentDate.getHours() == 12) {
    timeText = `今天 下午 ${currentDate.getHours()}:${currentDate.getMinutes()}`;
  } else if (currentDate.getHours() > 12) {
    timeText = `今天 下午 ${currentDate.getHours() - 12}:${currentDate.getMinutes()}`;
  }
  let picture = "morning.jpg"
  let hourStr = `${currentDate.getHours()}`
  if(currentDate.getHours() > 12){
    hourStr = `${currentDate.getHours() - 12}`
    picture = 'afternoon.jpg'
  }else if(currentDate.getHours() == 12){
    picture = 'afternoon.jpg'
  }
  const hourSplit = `12`.split("");
  let timeTotalLeft = 479;
  for (let i in hourSplit) {
    let picture = `./pic/time${hourSplit[i]}.jpg`
    timeTotalLeft = timeTotalLeft + i*15
    compositeArr.push({ input: picture, top: timeTop, left: timeTotalLeft });
  }
  timeTotalLeft = timeTotalLeft + 15
  compositeArr.push({ input: `./pic/timeColon.jpg`, top: timeTop, left: timeTotalLeft });

  timeTotalLeft = timeTotalLeft + 9
  let minuteStr = `${fillMinutes(currentDate.getMinutes())}`
  const minuteSplit = minuteStr.split("");
  for (let i in minuteSplit) {
    let picture = `./pic/time${minuteSplit[i]}.jpg`
    timeTotalLeft = timeTotalLeft + i*15
    compositeArr.push({ input: picture, top: timeTop, left: timeTotalLeft });
  }


  const fileName = crypto.randomBytes(20).toString('hex');

  const metadata = await sharp(`../${picture}`)
    .composite(compositeArr)
    .png()
    .toFile(`./${fileName}.jpg`);

  return `${fileName}.jpg`;
};

getMetadata("1234 1234 1234 123")