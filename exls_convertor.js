const XLSX = require('xlsx');
const fs = require('fs');

const workbook = XLSX.readFile('D:\\menus.xlsx');

// 첫 번째 시트 선택
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

// 데이터 변환
const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
// 첫 번째 행을 키로, 나머지 행을 값으로 변환
const keys = data[0];
const values = data.slice(1).map((row) => {
  const obj = {};
  row.slice(0, 4).forEach((cell, index) => {
    obj[keys[index]] = cell;
  });
  return obj;
});

const textData = values
  .map((row) => {
    const objString = JSON.stringify(row, null, 2).replace(/":/g, '": ');
    return objString;
  })
  .join(',\n');

// 텍스트 파일 저장
const textFilePath = 'output.txt';
fs.writeFileSync(textFilePath, textData, 'utf8');

console.log('텍스트 파일이 생성되었습니다.');

// 워크북 생성
const wb = XLSX.utils.book_new();

// 시트 생성
const ws = XLSX.utils.json_to_sheet(values);

// 워크북에 시트 추가
XLSX.utils.book_append_sheet(wb, ws, '신차이');

// 엑셀 파일 저장
const filePath = 'output.xlsx';
XLSX.writeFile(wb, filePath);

console.log('엑셀 파일이 생성되었습니다.');
