const XLSX = require('xlsx');

const workbook = XLSX.readFile('planilha.xlsx');
const sheetName = workbook.SheetNames[2];
const ws = workbook.Sheets[sheetName];

const sheetJson = XLSX.utils.sheet_to_json(ws).map((row) => {
	return Object.entries(row).reduce((acc, [key, value]) => {
		if (key === 'Nome') return { ...acc, [key]: 'Maça' };
		return { ...acc, [key]: value };
	}, {});
});

workbook.Sheets[sheetName] = XLSX.utils.json_to_sheet(sheetJson);
XLSX.writeFile(workbook, excelFile);
