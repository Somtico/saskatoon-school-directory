const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

async function convertExcelToJson() {
  try {
    // Read the Excel file
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(path.join(__dirname, "public-schools.xlsx"));

    // Get the first worksheet
    const worksheet = workbook.getWorksheet(1);

    // Convert to JSON
    const data = [];
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header row

      const school = {
        name: row.getCell(1).value,
        address: row.getCell(2).value,
        phone: row.getCell(3).value,
        email: row.getCell(4).value,
        contactPageUrl: row.getCell(5).value,
        type: row.getCell(6).value,
        language: row.getCell(7).value,
      };
      data.push(school);
    });

    // Save as JSON
    const jsonPath = path.join(__dirname, "public-schools.json");
    fs.writeFileSync(jsonPath, JSON.stringify(data, null, 2));
    console.log(`Data has been saved to: ${jsonPath}`);
  } catch (error) {
    console.error("Error converting Excel to JSON:", error);
  }
}

convertExcelToJson();
