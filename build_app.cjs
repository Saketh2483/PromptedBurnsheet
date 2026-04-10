// Build script to assemble App.jsx with embedded Excel data
const fs = require('fs');
const data = JSON.parse(fs.readFileSync('excel_mapped.json', 'utf8'));

// Minify the data - use short keys
const shortData = data.map(r => [
  r.esaId, r.esaDesc, r.vzTqId, r.vzTqDesc, r.poc, r.empId, r.name,
  r.location, r.country, r.actPct, r.skillSet, r.verizonLevel,
  r.classification, r.key, r.designation, r.serviceLine,
  r.timesheetHrs, r.rateInr, r.rateUsd, r.projectedRate,
  r.actualRate, r.variance, r.jan26, r.feb26, r.mar26
]);

const dataStr = JSON.stringify(shortData);
console.log('Compact array data size:', dataStr.length);
fs.writeFileSync('excel_arrays.json', dataStr);
console.log('Written excel_arrays.json');
