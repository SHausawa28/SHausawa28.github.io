const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});

// Define the route for handling form submissions
app.post('/submit', (req, res) => {
  const { teacherName, className, students } = req.body;

  // Create a new Excel workbook and worksheet
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Student Data');

  // Define the columns in the Excel sheet
  worksheet.columns = [
    { header: 'Teacher Name', key: 'teacherName' },
    { header: 'Class Name', key: 'className' },
    { header: 'Student Name', key: 'studentName' },
    { header: 'Roll Number', key: 'rollNumber' },
    { header: 'Score', key: 'score' },
  ];

  // Add the submitted data to the worksheet
  students.forEach((student) => {
    worksheet.addRow({
      teacherName,
      className,
      studentName: student.studentName,
      rollNumber: student.rollNumber,
      score: student.score,
    });
  });

  // Save the Excel file
  workbook.xlsx.writeFile('student_data.xlsx')
    .then(() => {
      console.log('Excel file saved successfully');
      res.status(201).json({ message: 'Student data saved successfully' });
    })
    .catch((err) => {
      console.error('Error saving Excel file:', err);
      res.status(500).json({ error: 'An error occurred while saving student data' });
    });
});
