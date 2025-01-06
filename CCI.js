// Helper to get the calculated CCI and CI values from CCI.xlsx
const getCalculatedCCIAndCI = async (outputData) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(cciExcelFilePath);
  const sheet = workbook.worksheets[0];

  for (const resource of outputData) {
    const { "GSO role": gsoRole } = resource;

    // Find the row in CCI sheet matching the GSO role
    sheet.eachRow((row) => {
      if (row.getCell(3).value === gsoRole) {  // Assuming GSO role is in column 3
        // Read the calculated CCI and CI values (now we extract only the result)
        const cci = row.getCell(7).value;  // Assuming CCI is in column 7
        const ci = row.getCell(8).value;   // Assuming CI is in column 8

        // Update the resource with only the result values for CCI and CI
        resource.CCI = typeof cci === "number" ? cci : parseFloat(cci.result || 0);
        resource.CI = typeof ci === "number" ? ci : parseFloat(ci.result || 0);
      }
    });
  }

  return outputData;
};
