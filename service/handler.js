const sharp = require('sharp');

module.exports.testAddCode = async function testAddCode(){
  const bufImage = await sharp('ImageTemplateBis.svg').grayscale().toFile('./result/test2.jpg');

  return { data: bufImage.toString('base64') };
};
