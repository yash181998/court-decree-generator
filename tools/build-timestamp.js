'use strict';

// Generates a build-info.js file with the current deployment timestamp.
// Run as part of: npm run build

const fs = require('fs');
const path = require('path');

const now = new Date();
const buildTime = now.toLocaleString('en-IN', {
  year: 'numeric',
  month: 'short',
  day: '2-digit',
  hour: '2-digit',
  minute: '2-digit',
  second: '2-digit',
  hour12: true,
  timeZone: 'Asia/Kolkata',
});

const jsContent = `// Generated at build time; do not edit manually.
export const BUILD_TIME = '${buildTime} IST';
`;

const target = path.join(__dirname, '..', 'js', 'build-info.js');
fs.writeFileSync(target, jsContent, 'utf8');
console.log(`Build timestamp: ${buildTime} IST`);
