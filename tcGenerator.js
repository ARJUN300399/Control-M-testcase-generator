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
  jobDefinitions.forEach((job) => {
    const testName = `TC${testCaseCounter}_${job.name} job validation`;
    const rowStart = testCaseCounter * 5 - 4;
    const rowEnd = rowStart + 4;

    const mergedColumns = [
      { index: 1, value: "L123" },
      { index: 2, value: "Draft" },
      { index: 3, value: "3-Medium" },
      { index: 4, value: testName },
      { index: 5, value: `To validate job ${job.name} for file watch for ${job.command}` },
      { index: 9, value: "SAS" },
      { index: 10, value: "Manual" },
      { index: 11, value: "No" },
      { index: 12, value: "Positive test" },
      { index: 13, value: "UUU" },
      { index: 14, value: "ooo" },
      { index: 15, value: "Other Needs" },
      { index: 16, value: "No" },
      { index: 17, value: "Unit testing" }
    ];

    const stepNames = ["Step 1", "Step 2", "Step 3", "Step 4", "Step 5"];
    const designStepsDesc = ["A", "B", "C", "D", "E"];
    const expectedResult = ["Abc", "Bce", "Cde", "Def", "Efg"];

    for (let rowNumber = rowStart; rowNumber <= rowEnd; rowNumber++) {
      const row = sheet.addRow([]);

      // Merge cells and set values for specified columns
      mergedColumns.forEach(({ index, value }) => {
        if (rowNumber === rowStart) {
          sheet.mergeCells(`A${rowStart}:A${rowEnd}`);
        }
        sheet.getCell(`${columns[index - 1]}${rowStart}`).value = value;
      });

      // Set values for Step Name, Design Steps Description, and Expected Result columns
      row.getCell("Step Name").value = stepNames[rowNumber - rowStart];
      row.getCell("Description (Design Steps)").value = designStepsDesc[rowNumber - rowStart];
      row.getCell("Expected Result (Design Steps)").value = expectedResult[rowNumber - rowStart];
    }

    testCaseCounter++;
  });

  // Save Excel file
  await workbook.xlsx.writeFile(outputExcelFile);
}
