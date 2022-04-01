const sharp = require('sharp');
const crypto = require('crypto');


module.exports.getMetadata = async function getMetadata(code) {
  const width = 500;
  const height = 300;
  const svg = `
<svg width="${width}" height="${height}" viewBox="0 0 ${height} ${height + 2}">
  <!--this rect should have rounded corners-->
  <rect x="-15" y="123" width="370" height="40" rx="1" fill="#34c759" />
  <text x="50%" y="50%" text-anchor="middle" dy="0.25em" fill="#e4f9ed" font-size="3em" font-family="-apple-system, BlinkMacSystemFont, sans-serif">${code}</text>
</svg>
`;
  const svg_buffer = Buffer.from(svg);

  // 日期
  const currentDate = new Date();
  let timeText = `今天 上午 ${currentDate.getHours()}:${currentDate.getMinutes()}`
  if(currentDate.getHours()==12){
    timeText = `今天 下午 ${currentDate.getHours()}:${currentDate.getMinutes()}`
  }else if(currentDate.getHours()>12){
    timeText = `今天 下午 ${currentDate.getHours()-12}:${currentDate.getMinutes()}`
  }
  const dateWidth = 300;
  const dateHeight = 300;
  const dateSvg = `
<svg width="${dateWidth}" height="${dateHeight}" viewBox="0 0 ${dateHeight} ${dateHeight + 2}">
  <!--this rect should have rounded corners-->
  <rect x="-15" y="123" width="370" height="40" rx="1" fill="#000000" />
  <text x="50%" y="50%" text-anchor="middle" dy="0.25em" fill="#8a898d" font-size="2em" font-family="-apple-system, BlinkMacSystemFont, sans-serif">${timeText}</text>
</svg>
`;
  const date_svg_buffer = Buffer.from(dateSvg);

  const fileName = crypto.randomBytes(20).toString('hex');
  const metadata = await sharp('S__37609519.jpg')
    .composite([
      {
        input: svg_buffer,
        top: 1365,
        left: 331,
      },
      {
        input: date_svg_buffer,
        top: 1300,
        left: 300,
      },
    ])
    .png()
    .toFile(`./result/${fileName}.jpg`);

    return `${fileName}.jpg`;
}

