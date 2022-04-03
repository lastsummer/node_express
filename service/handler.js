const sharp = require('sharp');

async function testAddCode(){
  const svg = `<?xml version="1.0" encoding="UTF-8"?>
<svg width="21" height="29" viewBox="0 0 21 37">
  <!--this rect should have rounded corners-->
  <rect x="-2" y="0" width="29" height="38" rx="1" fill="#34c759" />
  <text x="12" y="24" text-anchor="middle" dy="0.25em" fill="#e4f9ed" font-size="3.8em" font-family="-apple-system, BlinkMacSystemFont, sans-serif"> </text>
</svg>
`;
  const svg_buffer = Buffer.from(svg, 'utf8');

  const bufImage = await sharp(svg_buffer).png().toFile('./pic/numberSpace.jpg');

  return { data: bufImage.toString('base64') };
};

async function testAddTime(){
  const svg = `<?xml version="1.0" encoding="UTF-8"?>
<svg width="10" height="24" viewBox="0 0 10 32">
  <!--this rect should have rounded corners-->
  <rect x="-2" y="0" width="29" height="38" rx="1" fill="#000000" />
  <text x="3" y="19" text-anchor="middle" dy="0.25em" fill="#8a898d" font-size="2.8em" font-family="-apple-system, BlinkMacSystemFont, sans-serif">:</text>
</svg>
`;
  const svg_buffer = Buffer.from(svg, 'utf8');

  const bufImage = await sharp(svg_buffer).png().toFile('./pic/timeColon.jpg');

  return { data: bufImage.toString('base64') };
};
