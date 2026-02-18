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

  let records = classStudents.map((student) => {
    const markValue = req.body[`marks_${student.hallticket}`];
    const gradeValue = req.body[`grade_${student.hallticket}`];
    const isAbsent  = req.body[`absent_${student.hallticket}`];

    // 🟡 ABSENT HANDLING (highest priority)
    if (isAbsent) {
      return [
        teacherName || "",
        className || "",
        student.division || "",
        subjectName || "",
        examType || "",
        subjectName === "Drawing" || subjectName === "Supw" ? "" : minPassMark || "",
        subjectName === "Drawing" || subjectName === "Supw" ? "" : maxMarks || "",
        student.hallticket,
        student.name,
        "ABSENT",
        "Absent",
        timestamp
      ];
    }

    // 🎨 DRAWING / SUPW → Grades
    if (subjectName === "Drawing" || subjectName === "Supw") {
      if (!gradeValue) return null;
      return [
        teacherName || "",
        className || "",
        student.division || "",
        subjectName || "",
        examType || "",
        "",
        "",
        student.hallticket,
        student.name,
        gradeValue,
        "Grade",
        timestamp
      ];
    }

    // 🧮 MARKS SUBJECTS
    if (!markValue) return null;

    const result = (parseFloat(markValue) >= parseFloat(minPassMark))
      ? "Pass"
      : "Fail";

    return [
      teacherName || "",
      className || "",
      student.division || "",
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
  }).filter(Boolean);

  if (records.length === 0) {
    return res.send("No marks / grades / absentees entered!");
  }

  const sheetName = `${className}-${subjectName}`;

  try {
    const client = await auth.getClient();
    const sheets = google.sheets({ version: "v4", auth: client });

    const sheetMeta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    const existingSheets = sheetMeta.data.sheets.map(s => s.properties.title);

    if (!existingSheets.includes(sheetName)) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: {
          requests: [{ addSheet: { properties: { title: sheetName } } }]
        }
      });

      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `${sheetName}!A1:L1`,
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [[
            "Teacher",
            "Class",
            "Division",
            "Subject",
            "Exam Type",
            "Min Pass",
            "Max Marks",
            "Hallticket",
            "Student Name",
            "Marks/Grade",
            "Result",
            "Timestamp"
          ]]
        }
      });
    }

    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: `${sheetName}!A:L`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: records }
    });

    // 🔴 Coloring FAIL only (skip Drawing & SUPW)
    if (subjectName !== "Drawing" && subjectName !== "Supw") {
  const freshMeta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const targetSheet = freshMeta.data.sheets.find(
    s => s.properties.title === sheetName
  );
  const sheetId = targetSheet.properties.sheetId;

  // Get entire Result column (K)
  const resultData = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${sheetName}!K2:K`,
  });

  const results = resultData.data.values || [];

  const requests = results.map((row, i) => {
    const value = row[0] ? row[0].toString().trim().toLowerCase() : "";

    let color = { red: 1, green: 1, blue: 1 }; // default white

    if (value === "fail") {
      // 🔴 Light red
      color = { red: 1, green: 0.8, blue: 0.8 };
    } else if (value === "absent") {
      // 🟡 Light yellow
      color = { red: 1, green: 1, blue: 0.6 };
    }

    return {
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: i + 1,   // Skip header
          endRowIndex: i + 2,
          startColumnIndex: 0,    // Start from column A
          endColumnIndex: 12      // Until column L
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: color
          }
        },
        fields: "userEnteredFormat.backgroundColor"
      }
    };
  });

  if (requests.length > 0) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests }
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
