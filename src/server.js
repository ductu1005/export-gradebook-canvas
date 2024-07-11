const express = require("express");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const archiver = require("archiver");
// const fetch = require('node-fetch');

const app = express();
const port = xxx;
const token = "xxx";

app.use(express.json());
app.use(express.static(path.join(__dirname, "..", "public")));

function roundToOneDecimalPlace(value) {
  if (value == null) return null;
  return parseFloat(value.toFixed(1));
}

async function fetchFromCanvas(apiUrl) {
  try {
    const response = await fetch(apiUrl, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }
    return await response.json();
  } catch (error) {
    console.error("Error fetching data:", error);
    throw error;
  }
}

async function fetchPaginatedData(apiUrl) {
  let allData = [];
  let page = 1;
  while (true) {
    let apiUrlNew = apiUrl;
    if (apiUrlNew.endsWith('?')) {
      apiUrlNew += `page=${page}`;
    } else {
      apiUrlNew += `&page=${page}`;
    }
    console.log(apiUrlNew);
    const data = await fetchFromCanvas(`${apiUrlNew}`);
    if (data.length === 0 || (Array.isArray(data.enrollment_terms) && data.enrollment_terms.length === 0)) {
        break;
    }
    if (Array.isArray(data.enrollment_terms)) {
        allData = allData.concat(data.enrollment_terms);
    } else {
        allData = allData.concat(data);
    }
    page++;
  }
  // console.log("data: ", allData);
  return allData;
}

async function fetchAssignments(courseId, assignmentId) {
  const apiUrl = `${ip}/api/v1/courses/${courseId}/assignments/${assignmentId}/submissions?`;
  return fetchPaginatedData(apiUrl);
}

async function fetchStudents(courseId) {
  const apiUrl = `${ip}/api/v1/courses/${courseId}/users?enrollment_type[]=student`;
  return fetchPaginatedData(apiUrl);
}

// api canvas get all terms
app.get("/api/terms", async (req, res) => {
    try {
      const courses = await fetchPaginatedData(`${ip}/api/v1/accounts/1/terms?`);
      res.json(courses);
    } catch (error) {
      res.status(500).json({ error: "Error fetching data" });
    }
});

// api canvas get all course
app.get("/api/courses/:termId", async (req, res) => {
  try {
    const { termId } = req.params;
    const courses = await fetchPaginatedData(`${ip}/api/v1/accounts/1/courses?enrollment_term_id=${termId}`);
    res.json(courses);
  } catch (error) {
    res.status(500).json({ error: "Error fetching data" });
  }
});

// api canvas get assignments in course 
app.get("/api/courses/:courseId/assignments", async (req, res) => {
  const { courseId } = req.params;
  try {
    const assignments = await fetchPaginatedData(`${ip}/api/v1/courses/${courseId}/assignments?`);
    res.json(assignments);
  } catch (error) {
    res.status(500).json({ error: "Error fetching data" });
  }
});

app.post("/export-excel", async (req, res) => {
  const { courseId, scoreType, componentScore, finalScore } = req.body;
  const templateFilePath = path.join(__dirname, "..", "template.xlsx");
  const outputDir = path.join(__dirname, "..", "uploads", "output");
  const zipFileName = `output_${Date.now()}.zip`;

  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  try {
    const [componentScoreData, finalScoreData, studentData] = await Promise.all([
      fetchAssignments(courseId, componentScore),
      fetchAssignments(courseId, finalScore),
      fetchStudents(courseId)
    ]);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templateFilePath);
    const worksheet = workbook.getWorksheet(1);

    let startRow = 12;
    const numberOfRowsToInsert = studentData.length;
    worksheet.getCell("A5").value = `${worksheet.getCell("A5").value} ${scoreType}`;
    worksheet.getCell("A13").value = `${worksheet.getCell("A13").value} ${numberOfRowsToInsert} sinh viên`;
    worksheet.spliceRows(startRow, 0, ...Array(numberOfRowsToInsert).fill([]));

    studentData.forEach((student, index) => {
      const rowNumber = startRow + index;
      worksheet.getCell(`A${rowNumber}`).value = index + 1;
      worksheet.getCell(`B${rowNumber}`).value = student.sis_user_id;
      worksheet.getCell(`C${rowNumber}`).value = student.name;

      const componentScoreEntry = componentScoreData.find(entry => entry.user_id === student.id);
      const componentScoreValue = componentScoreEntry ? componentScoreEntry.score : null;

      const finalScoreEntry = finalScoreData.find(entry => entry.user_id === student.id);
      const finalScoreValue = finalScoreEntry ? finalScoreEntry.score : null;

      worksheet.getCell(`F${rowNumber}`).value = componentScoreValue;
      worksheet.getCell(`G${rowNumber}`).value = finalScoreValue;
      worksheet.getCell(`H${rowNumber}`).value = roundToOneDecimalPlace(
        worksheet.getCell(`F${rowNumber}`).value * worksheet.getCell("G7").value +
        worksheet.getCell(`G${rowNumber}`).value * worksheet.getCell("G8").value
      );

      ["A", "B", "C", "D", "E", "F", "G", "H"].forEach(col => {
        worksheet.getCell(`${col}${rowNumber}`).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });
    });

    const outputFilePath = path.join(outputDir, `output_${Date.now()}.xlsx`);
    await workbook.xlsx.writeFile(outputFilePath);

    const zipFilePath = path.join(__dirname, "..", "uploads", zipFileName);
    const output = fs.createWriteStream(zipFilePath);
    const archive = archiver("zip", { zlib: { level: 9 } });

    output.on("close", () => {
      res.setHeader("Content-Type", "application/zip");
      res.setHeader("Content-Disposition", `attachment; filename=${zipFileName}`);
      res.download(zipFilePath, () => {
        fs.unlinkSync(outputFilePath);
        fs.unlinkSync(zipFilePath);
      });
    });

    archive.pipe(output);
    archive.file(outputFilePath, { name: path.basename(outputFilePath) });
    await archive.finalize();
  } catch (err) {
    console.error("Error converting to Excel:", err);
    res.status(500).json({ success: false, error: "Error converting to Excel" });
  }
});

app.listen(port, () => {
  console.log(`Server running on http://10.10.0.128:${port}`);
});
