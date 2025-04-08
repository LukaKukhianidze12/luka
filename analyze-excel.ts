const XLSX = require('xlsx');
import * as fs from 'fs';
import * as path from 'path';

const filePath = path.resolve('./attached_assets/გადაწყობილი.xlsx');

try {
  // Read the Excel file
  const workbook = XLSX.readFile(filePath);
  
  // Get the first worksheet
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  
  // Convert the worksheet to JSON
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
  
  // Print headers to understand the structure
  console.log("Headers:", jsonData[0]);
  console.log("First row:", jsonData[1]);
  console.log("Total rows:", jsonData.length);
  
  // Analyze the Excel structure
  console.log("\nExcel Structure Analysis:");
  let startRow = -1;
  
  // Find where expert data starts (mentioned as A3 in the description)
  for (let i = 0; i < jsonData.length; i++) {
    const row = jsonData[i] as any[];
    if (row && row[0] && typeof row[0] === 'string' && row[0].match(/^[A-Za-z]/)) {
      console.log(`Expert data starts at row ${i+1} (index ${i}):`, row);
      startRow = i;
      break;
    }
  }
  
  if (startRow >= 0) {
    // Identify columns
    const sampleRow = jsonData[startRow] as any[];
    console.log("\nColumn Analysis:");
    for (let i = 0; i < sampleRow.length; i++) {
      if (sampleRow[i] !== null) {
        console.log(`Column ${String.fromCharCode(65 + i)} (${i})`, sampleRow[i]);
      }
    }
    
    // Get a few sample data rows
    console.log("\nSample Data:");
    for (let i = startRow; i < Math.min(startRow + 5, jsonData.length); i++) {
      console.log(`Row ${i+1}:`, jsonData[i] as any[]);
    }
  }
  
} catch (error) {
  console.error("Error analyzing Excel file:", error);
}