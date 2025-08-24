require("dotenv").config();
const express = require("express");
const bodyParser = require("body-parser");
const fs = require("fs");
const { google } = require("googleapis");
const path = require("path");  // for image loading

const app = express();
app.use(bodyParser.urlencoded({ extended: true }));
app.set("view engine", "ejs");

app.use(express.static(path.join(__dirname, "public")));

// Load students JSON
const students = JSON.parse(fs.readFileSync("students.json", "utf-8"));

// Extract unique classes for dropdown
const classList = [...new Set(students.map(s => s.className))];

// Google Sheets setup
const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const SHEET_ID = "1_KyTkxBux30HV4Sf4vd49cfXRUdbOPY4SZWEZhEAGH0";

// Step 1: Select class
app.get("/", (req, res) => {
  res.render("selectClass", { classList });
});

// Step 2: Show marks form for chosen class
app.get("/class/:className", (req, res) => {
  const selectedClass = req.params.className;
  const classStudents = students.filter(s => s.className === selectedClass);

  if (classStudents.length === 0) {
    return res.send("No students found for this class.");
  }

  res.render("form", { students: classStudents, className: selectedClass });
});

// Step 3: Submit marks
app.post("/submit", async (req, res) => {
  const { teacherName, subjectName, className, examType, minPassMark, maxMarks } = req.body;
  const classStudents = students.filter(s => s.className === className);

  const timestamp = new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" });

  // ✅ Build records for either Marks (default) or Grades (Drawing)
  let records = classStudents.map((student) => {
    const markValue = req.body[`marks_${student.hallticket}`];
    const gradeValue = req.body[`grade_${student.hallticket}`];

    if (subjectName === "Drawing" | subjectName ===  "Supw") {
      if (!gradeValue) return null; // skip empty
      return [
        teacherName || "",
        className || "",
        subjectName || "",
        examType || "",
        "", // Min Pass not applicable
        "", // Max Marks not applicable
        student.hallticket,
        student.name,
        gradeValue,   // store grade instead of marks
        "Grade",      // explicitly show it's a grade
        timestamp
      ];
    } else {
      if (!markValue) return null; // skip empty
      const result = (parseFloat(markValue) >= parseFloat(minPassMark)) ? "Pass" : "Fail";
      return [
        teacherName || "",
        className || "",
        subjectName || "",
        examType || "",
        minPassMark || "",
        maxMarks || "",
        student.hallticket,
        student.name,
        markValue,
        result,
        timestamp
      ];
    }
  }).filter(Boolean);

  // ✅ Ensure at least one record exists
  if (records.length === 0) {
    return res.send("No marks/grades entered!");
  }

  // ✅ Sheet name = className-subjectName
  const sheetName = `${className}-${subjectName}`;

  try {
    const client = await auth.getClient();
    const sheets = google.sheets({ version: "v4", auth: client });

    // Check if worksheet exists
    const sheetMeta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    const existingSheets = sheetMeta.data.sheets.map(s => s.properties.title);

    if (!existingSheets.includes(sheetName)) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: {
          requests: [{ addSheet: { properties: { title: sheetName } } }]
        }
      });

      // ✅ Add headers
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `${sheetName}!A1:K1`,
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [[
            "Teacher",
            "Class",
            "Subject",
            "Exam Type",
            "Min Pass",
            "Max Marks",
            "Hallticket",
            "Student Name",
            "Marks/Grade",
            "Result",
            "Timestamp"
          ]],
        },
      });
    }

    // ✅ Append data
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: `${sheetName}!A:K`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: records },
    });

    // ✅ Color "Fail" rows only when it's not Drawing
    if (subjectName !== "Drawing" | subjectName !== "Supw") {
      const freshMeta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
      const targetSheet = freshMeta.data.sheets.find(s => s.properties.title === sheetName);
      const sheetId = targetSheet.properties.sheetId;

      const failRequests = [];
      records.forEach((row, i) => {
        if (row[9] === "Fail") { // column J = index 9
          failRequests.push({
            repeatCell: {
              range: {
                sheetId,
                startRowIndex: i + 1,  // skip header
                endRowIndex: i + 2,
                startColumnIndex: 9,
                endColumnIndex: 10
              },
              cell: {
                userEnteredFormat: {
                  backgroundColor: { red: 1, green: 0.8, blue: 0.8 }
                }
              },
              fields: "userEnteredFormat.backgroundColor"
            }
          });
        }
      });

      if (failRequests.length > 0) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: SHEET_ID,
          requestBody: { requests: failRequests }
        });
      }
    }

    res.render("success", { message: `Data for ${sheetName} submitted successfully!` });

  } catch (err) {
    console.error(err);
    res.send("Error saving data to Google Sheets");
  }
});
const PORT = process.env.PORT || 3000;

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
