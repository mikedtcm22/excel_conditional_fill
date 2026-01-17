// Simple icon generator for placeholder development
// Creates basic colored squares as placeholder icons

const fs = require('fs');
const path = require('path');

// Base64 encoded 1x1 blue PNG
const bluePNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAEBgIApD5fRAAAAABJRU5ErkJggg==', 'base64');

const assetsDir = path.join(__dirname, '../assets');

// Create placeholder PNGs at different sizes (all using the same 1x1 placeholder for now)
const sizes = [16, 32, 64, 80];

if (!fs.existsSync(assetsDir)) {
  fs.mkdirSync(assetsDir, { recursive: true });
}

sizes.forEach(size => {
  const filename = `icon-${size}.png`;
  const filepath = path.join(assetsDir, filename);
  fs.writeFileSync(filepath, bluePNG);
  console.log(`Created ${filename}`);
});

console.log('Placeholder icons generated successfully');
