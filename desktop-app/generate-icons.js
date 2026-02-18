/**
 * Generate app icons from SVG for Windows (.ico) and Mac (.icns)
 * Requires: npm install sharp png-to-ico
 */

const fs = require('fs');
const path = require('path');

async function generateIcons() {
  console.log('Generating app icons from SVG...');

  // Try to use sharp for PNG generation
  let sharp;
  try {
    sharp = require('sharp');
  } catch (e) {
    console.log('sharp not available, skipping icon generation');
    return;
  }

  const svgPath = path.join(__dirname, 'icon.svg');

  if (!fs.existsSync(svgPath)) {
    console.log('icon.svg not found, skipping icon generation');
    return;
  }

  const svgBuffer = fs.readFileSync(svgPath);

  // Icon sizes needed
  const sizes = [16, 32, 48, 64, 128, 256, 512];

  // Create build directory
  const buildDir = path.join(__dirname, 'build');
  if (!fs.existsSync(buildDir)) {
    fs.mkdirSync(buildDir);
  }

  // Generate PNGs for each size
  console.log('Generating PNG files...');
  const pngBuffers = {};
  const pngPaths = [];

  for (const size of sizes) {
    try {
      const pngBuffer = await sharp(svgBuffer)
        .resize(size, size)
        .png()
        .toBuffer();

      const pngPath = path.join(buildDir, `icon_${size}x${size}.png`);
      fs.writeFileSync(pngPath, pngBuffer);
      pngBuffers[size] = pngBuffer;
      pngPaths.push(pngPath);
      console.log(`  Created ${size}x${size} PNG`);
    } catch (e) {
      console.log(`  Failed to create ${size}x${size}: ${e.message}`);
    }
  }

  // Generate ICO for Windows using png-to-ico
  try {
    const pngToIco = require('png-to-ico');

    // Use 256, 48, 32, 16 for ICO
    const icoPngs = [256, 48, 32, 16]
      .filter(s => pngBuffers[s])
      .map(s => path.join(buildDir, `icon_${s}x${s}.png`));

    if (icoPngs.length > 0) {
      const icoBuffer = await pngToIco(icoPngs);
      fs.writeFileSync(path.join(__dirname, 'icon.ico'), icoBuffer);
      console.log('Created icon.ico');
    }
  } catch (e) {
    console.log('png-to-ico failed, creating simple ICO:', e.message);
    // Fallback: create simple ICO from 256px PNG
    if (pngBuffers[256]) {
      const ico = createSimpleIco(pngBuffers[256]);
      fs.writeFileSync(path.join(__dirname, 'icon.ico'), ico);
      console.log('Created icon.ico (simple fallback)');
    }
  }

  // Copy 256px PNG as icon.png for electron-builder
  if (pngBuffers[256]) {
    fs.writeFileSync(path.join(__dirname, 'icon.png'), pngBuffers[256]);
    console.log('Created icon.png (256x256)');
  }

  // For Mac, create iconset folder structure
  const iconsetDir = path.join(buildDir, 'icon.iconset');
  if (!fs.existsSync(iconsetDir)) {
    fs.mkdirSync(iconsetDir);
  }

  // Mac iconset naming convention
  const macSizes = [
    { size: 16, name: 'icon_16x16.png' },
    { size: 32, name: 'icon_16x16@2x.png' },
    { size: 32, name: 'icon_32x32.png' },
    { size: 64, name: 'icon_32x32@2x.png' },
    { size: 128, name: 'icon_128x128.png' },
    { size: 256, name: 'icon_128x128@2x.png' },
    { size: 256, name: 'icon_256x256.png' },
    { size: 512, name: 'icon_256x256@2x.png' },
    { size: 512, name: 'icon_512x512.png' },
  ];

  for (const { size, name } of macSizes) {
    if (pngBuffers[size]) {
      fs.writeFileSync(path.join(iconsetDir, name), pngBuffers[size]);
    }
  }
  console.log('Created Mac iconset folder');

  // On Mac, try to generate ICNS using iconutil
  if (process.platform === 'darwin') {
    const { execSync } = require('child_process');
    try {
      execSync(`iconutil -c icns "${iconsetDir}" -o "${path.join(__dirname, 'icon.icns')}"`);
      console.log('Created icon.icns using iconutil');
    } catch (e) {
      console.log('Could not create ICNS with iconutil:', e.message);
    }
  }

  console.log('Icon generation complete!');
}

function createSimpleIco(pngBuffer) {
  // Very basic ICO creation - wraps PNG in ICO container
  const imageSize = pngBuffer.length;

  // ICO Header (6 bytes)
  const header = Buffer.alloc(6);
  header.writeUInt16LE(0, 0);     // Reserved
  header.writeUInt16LE(1, 2);     // Type (1 = ICO)
  header.writeUInt16LE(1, 4);     // Number of images

  // Directory entry (16 bytes)
  const entry = Buffer.alloc(16);
  entry.writeUInt8(0, 0);          // Width (0 = 256)
  entry.writeUInt8(0, 1);          // Height (0 = 256)
  entry.writeUInt8(0, 2);          // Color palette
  entry.writeUInt8(0, 3);          // Reserved
  entry.writeUInt16LE(1, 4);       // Color planes
  entry.writeUInt16LE(32, 6);      // Bits per pixel
  entry.writeUInt32LE(imageSize, 8);  // Image size
  entry.writeUInt32LE(22, 12);     // Offset to image data (6 + 16 = 22)

  return Buffer.concat([header, entry, pngBuffer]);
}

generateIcons().catch(console.error);
