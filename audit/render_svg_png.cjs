// Render vectors directly at print resolution; do not upscale the VB6 screenshot.
const sharp = require('sharp');
const fs = require('fs');
const [input, output] = process.argv.slice(2);
if (!input || !output) throw new Error('Usage: node render_svg_png.cjs input.svg output.png');
const svg = fs.readFileSync(input, 'utf8');
const cm = key => Number(svg.match(new RegExp(`${key}="([\\d.]+)cm"`))[1]);
const width = Math.round(cm('width') / 2.54 * 600);
const height = Math.round(cm('height') / 2.54 * 600);
sharp(input, {density: 600})
  .resize(width, height, {fit: 'fill'})
  .withMetadata({density: 600})
  .png()
  .toFile(output)
  .then(() => sharp(output).metadata())
  .then(m => console.log(JSON.stringify({width: m.width, height: m.height, density: m.density})))
  .catch(e => { console.error(e); process.exitCode = 1; });
