const xlsx = require('xlsx');
const fs = require('fs');

const workbook = xlsx.readFile('ListadodeProductos.xlsx');
const sheet_name_list = workbook.SheetNames;
const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

fs.writeFileSync('products.json', JSON.stringify(data, null, 2));
console.log(`Exported ${data.length} products to products.json`);
