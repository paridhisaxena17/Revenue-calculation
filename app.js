const express = require("express");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const app = express();
app.use(express.json());

// File paths
const excelFilePath = path.join('C:\\Users\\admin\\paridhi\\excel sheet', 'Employee Master.xlsx');
const mmeExcelFilePath = path.join('C:\\Users\\admin\\paridhi\\excel sheet', 'MME.xlsx');
const lcrExcelFilePath = path.join('C:\\Users\\admin\\paridhi\\excel sheet', 'LCR calculation.xlsx');
const outputFilePath = path.join('C:\\Users\\admin\\paridhi\\excel sheet', 'output.json');

// Helper to convert Excel to JSON using ExcelJS
const convertExcelToJson = async (filePath) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.worksheets[0];

  const jsonData = [];
  const headers = sheet.getRow(1).values.slice(1); // Skip the first empty column

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // Skip header row
    const rowData = {};
    headers.forEach((header, index) => {
      rowData[header] = row.getCell(index + 1).value;
    });
    jsonData.push(rowData);
  });

  return jsonData;
};

// Helper to process LCR using Excel sheet calculation
const processLCR = async (filteredData) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(lcrExcelFilePath);
  const sheet = workbook.worksheets[0];

  for (const resource of filteredData) {
    const { Level, "Bill code": billCode, "Unloaded Amount": unloadedAmt } = resource;

    // Find the row in LCR sheet matching the Level
    let rowToUpdate;
    sheet.eachRow((row, rowNumber) => {
      if (row.getCell(1).value === Level) {
        rowToUpdate = row;
      }
    });

    if (rowToUpdate) {
      // Update Bill Code and Unloaded Amount in the Excel sheet
      rowToUpdate.getCell(2).value = billCode; // Column B: Bill Code
      rowToUpdate.getCell(3).value = unloadedAmt; // Column C: Unloaded Amount

      // Save changes to Excel file
      await workbook.xlsx.writeFile(lcrExcelFilePath);

      // Reload the updated Excel file to get the recalculated LCR value
      const updatedWorkbook = new ExcelJS.Workbook();
      await updatedWorkbook.xlsx.readFile(lcrExcelFilePath);
      const updatedSheet = updatedWorkbook.worksheets[0];

      updatedSheet.eachRow((row) => {
        if (row.getCell(1).value === Level) {
          const lcrValue = row.getCell(6).value; // Column F: LCR
          resource.LCR = typeof lcrValue === "number" ? lcrValue : parseFloat(lcrValue.result || 0);
        }
      });
    }
  }

  return filteredData;
};

// Endpoint to filter data and calculate LCR
app.post("/filter-and-calculate", async (req, res) => {
  try {
    // Load Employee Master data
    const employeeData = await convertExcelToJson(excelFilePath);

    // Load MME data
    const mmeData = await convertExcelToJson(mmeExcelFilePath);

    // Filter Employee Master data based on request
    const filters = req.body;
    let filteredData = employeeData;

    Object.keys(filters).forEach((key) => {
      filteredData = filteredData.filter((row) => row[key] === filters[key]);
    });

    // Merge MME data into filtered Employee Master data
    filteredData = filteredData.map((resource) => {
      const mmeRow = mmeData.find((mme) => mme["Enterprise ID"] === resource["Enterprise ID"]);
      if (mmeRow) {
        resource["Bill code"] = mmeRow["Bill code"];
        resource["Unloaded Amount"] = mmeRow["Unloaded Amount"];
      }
      return resource;
    });

    // Process LCR for the filtered data
    const finalData = await processLCR(filteredData);

    // Save the result to output.json
    fs.writeFileSync(outputFilePath, JSON.stringify(finalData, null, 2));

    // Send the final data as a response
    res.json(finalData);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Start the server
const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
