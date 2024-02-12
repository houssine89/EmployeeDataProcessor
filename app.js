

const fs = require('fs');
const xlsx = require('xlsx');

// Replace 'example.xlsx' with the path to your Excel file
const excelFilePath = './exel.xlsx';

// Check if the file exists before attempting to read it

  try {
    // Read the Excel file
    const workbook = xlsx.readFile(excelFilePath);

    // Assume the first sheet is the one you want to work with
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Convert the sheet to JSON
    const data = xlsx.utils.sheet_to_json(worksheet, { header: 'A' });

    //console.log('Excel data:', jsonData);

    const newdata = data.map(employee => {
        const annualSalary = employee.AnnualSalary;
        let bonusPercentage, bonusAmount;
        if (annualSalary < 50000) {
            bonusPercentage = 5;
        } else if (annualSalary >= 50000 && annualSalary <= 100000) {
            bonusPercentage = 7;
        } else {
            bonusPercentage = 10;
        }
        bonusAmount = (annualSalary * bonusPercentage) / 100;
        return {
            ...employee,
            BonusPercentage: bonusPercentage,
            BonusAmount: bonusAmount
        };
    });

    const newWorkbook = xlsx.utils.book_new();
    const newWorksheet = xlsx.utils.json_to_sheet(newdata);
    xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);

    // Write the new workbook to a new Excel file
    const outputFilePath = 'newExel.xlsx';
    xlsx.writeFile(newWorkbook, outputFilePath);

    console.log('Bonus calculation completed. Results saved to:', outputFilePath);

  } catch (error) {
    console.error('Error reading Excel file:', error.message);
  }
/*
} else {
  console.error('Error: The specified file does not exist.');
}











       /* const newdata = data.map(employee => {
        const annualSalary =employee.AnnualSalary
        let bonusPercentage, bonusAmount;
        if (annualSalary < 50000) {
            bonusPercentage = 5;
        } else if (annualSalary >= 50000 && annualSalary <= 100000) {
            bonusPercentage = 7;
        } else {
            bonusPercentage = 10;
        }
        bonusAmount = (annualSalary * bonusPercentage) / 100;
        return {
            ...employee,
            BonusPercentage: bonusPercentage,
            BonusAmount: bonusAmount.toFixed(2)
        };
    });
*/

