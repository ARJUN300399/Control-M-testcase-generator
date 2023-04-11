const fs = require("fs");
const xml2js = require("xml2js");
const ExcelJS = require("exceljs");

(async () => {
  const jobDefinitionsFile = "path/to/your/job_definitions.xml";
  const jobDefinitions = await parseJobDefinitions(jobDefinitionsFile);

  const outputExcelFile = "path/to/your/testing_report.xlsx";
  if (fs.existsSync(outputExcelFile)) {
    fs.unlinkSync(outputExcelFile);
  }
  await generateExcelReport(jobDefinitions, outputExcelFile);
})();

async function parseJobDefinitions(jobDefinitionsFile) {
  const xmlContent = fs.readFileSync(jobDefinitionsFile, "utf8");
  const parser = new xml2js.Parser();

  const parsedXml = await parser.parseStringPromise(xmlContent);
  const jobNodes = parsedXml.DEFTABLE.FOLDER.flatMap(folder => folder.JOB);

  const jobDefinitions = jobNodes.map(jobNode => {
    const commandWithQuotes = jobNode.$.CHOLINE;
    const commandWithoutQuotes = commandWithQuotes.replace(/&quot;/g, "").trim();
    const commandParts = commandWithoutQuotes.split(" ");
    const filePath = commandParts.shift();
    const fileName = filePath.substring(filePath.lastIndexOf("\\") + 1);

    return {
      jobNumber: jobNode.$.JOBSN,
      name: jobNode.$.JOBNAME,
      command: commandParts.join(" "),
      host: jobNode.$.NODEID,
      user: jobNode.$.RUN_AS,
      filePath: filePath,
      fileName: fileName
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
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFBFBFBF" }
    };
  });

  // Add test case rows
  let testCaseCounter = 1;
  jobDefinitions.forEach((job) => {
    const testName = `TC${testCaseCounter}_${job.name} job validation`;
    const rowStart = testCaseCounter * 5;
    const rowEnd = rowStart + 4;

    const stepNames = ["Step 1", "Step 2", "Step 3", "Step 4", "Step 5"];
    const designStepDescriptions = ["Auto", "Mation", "Noida", "Delhi", "Alwar"];
    const expectedResultDescriptions = ["Hello", "All", "Project", "Wale", "Efg"];

    for (let rowNumber = rowStart; rowNumber <= rowEnd; rowNumber++) {
      const row = sheet.addRow([]);

      // Merge cells and set values for specified columns
      if (rowNumber === rowStart) {
        const mergedColumns = [
          { col: "A", value: "L123" },
          { col: "B", value: "Draft" },
          { col: "C", value: "3-Medium" },
          { col: "D", value: testName },
          { col: "E", value: `To validate job ${job.name} for file watch for ${job.command}` },
          { col: "I", value: "SAS" },
          { col: "J", value: "Manual" },
          { col: "K", value: "No" },
          { col: "L", value: "Positive test" },
          { col: "M", value: "UUU" },
          { col: "N", value: "ooo" },
          { col: "O", value: "Other Needs" },
          { col: "P", value: "No" },
          { col: "Q", value: "Unit testing" },
        ];

        mergedColumns.forEach(({ col, value }) => {
          sheet.mergeCells(`${col}${rowStart}:${col}${rowEnd}`);
          sheet.getCell(`${col}${rowStart}`).value = value;
        });
      }

      // Set values for Step Name, Description (Design Steps), and Expected Result (Design Steps) columns
      const stepIndex = rowNumber - rowStart;
      row.getCell("F").value = stepNames[stepIndex];
      row.getCell("G").value = designStepDescriptions[stepIndex];
      row.getCell("H").value = expectedResultDescriptions[stepIndex];
    }

    testCaseCounter++;
  });

  // Save Excel file
  await workbook.xlsx.writeFile(outputExcelFile);
}
