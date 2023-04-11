const fs = require("fs");
const xml2js = require("xml2js");
const ExcelJS = require("exceljs");

(async () => {
  const jobDefinitionsFile = "path/to/your/job_definitions.xml";
  const jobDefinitions = await parseJobDefinitions(jobDefinitionsFile);

  const outputExcelFile = "path/to/your/testing_report.xlsx";
  await generateExcelReport(jobDefinitions, outputExcelFile);
})();

async function parseJobDefinitions(jobDefinitionsFile) {
  const xmlContent = fs.readFileSync(jobDefinitionsFile, "utf8");
  const parser = new xml2js.Parser();

  const parsedXml = await parser.parseStringPromise(xmlContent);
  const jobNodes = parsedXml.DEFTABLE.FOLDER.flatMap(folder => folder.JOB);

  const jobDefinitions = jobNodes.map(jobNode => {
    return {
      jobNumber: jobNode.$.JOBSN,
      name: jobNode.$.JOBNAME,
      command: jobNode.$.CHOLINE,
      host: jobNode.$.NODEID,
      user: jobNode.$.RUN_AS
    };
  });

  return jobDefinitions;
}

function testConnectivity(host) {
  // Your connectivity testing code here
}

function testServiceAccount(user, command, host) {
  // Your service account testing code here
}

async function generateExcelReport(jobDefinitions, outputExcelFile) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Jobs");

  // Column names
  const columns = [
    "LID",
    "Status",
    "Test Priority",
    "Test Name",
    "Description",
    "Step Name",
    "Description (Design Steps)",
    "Expected Result (Design Steps)",
    "Module",
    "Type",
    "Regression",
    "Test Case Type",
    "Subject",
    "Application",
    "Remarks",
    "Automation",
    "Test Phase"
  ];

  // Header row
  sheet.addRow(columns).eachCell((cell, colNumber) => {
    cell.font = { bold: true, name: "Calibri", size: 11 };
  });

  // Add test case rows
  let testCaseCounter = 1;
  jobDefinitions.forEach(job => {
    const testName = `TC${testCaseCounter}_${job.name} job validation`;
    const rowStart = testCaseCounter * 5;
    const rowEnd = rowStart + 4;

    for (let rowNumber = rowStart; rowNumber <= rowEnd; rowNumber++) {
      const row = sheet.addRow([]);

      // Merge cells and set values for LID, Status, Test Priority, and Test Name columns
      if (rowNumber === rowStart) {
        sheet.mergeCells(`A${rowStart}:A${rowEnd}`);
        sheet.getCell(`A${rowStart}`).value = "L123";

        sheet.mergeCells(`B${rowStart}:B${rowEnd}`);
        sheet.getCell(`B${rowStart}`).value = "Draft";

        sheet.mergeCells(`C${rowStart}:C${rowEnd}`);
        sheet.getCell(`C${rowStart}`).value = "3-Medium";

        sheet.mergeCells(`D${rowStart}:D${rowEnd}`);
        sheet.getCell(`D${rowStart}`).value = testName;
      }
    }

    testCaseCounter++;
  });

  // Save Excel file
  await workbook.xlsx.writeFile(outputExcelFile);
}
