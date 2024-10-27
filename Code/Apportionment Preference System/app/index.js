const express = require("express");
const cors = require("cors");
const fs = require("fs");
const csv = require("csv-parser");
const path = require("path");
const XLSX = require("xlsx"); 

const app = express();
const port = 3001;
const dataFilePath = path.join(__dirname, "public",'data.json');
app.use(cors());
app.use(express.urlencoded({ extended: false }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'pages')));

const readDataFromFile = () => {
  const data = fs.readFileSync(dataFilePath);
  return JSON.parse(data);
};
const writeDataToFile = (data) => {
  fs.writeFileSync(dataFilePath, JSON.stringify(data, null, 2));
};
app.post('/api/data', (req, res) => {
  const { newQuantity, newScore } = req.body;
  const data = readDataFromFile();
  if (newQuantity !== undefined) {
      data.quantity = newQuantity;
  }

  if (newScore !== undefined) {
      data.score = newScore;
  }
  writeDataToFile(data);
  res.json(data);
});
app.get('/', (req, res) => {
  res.redirect('/1.html');
});

app.get("/api/projects", (req, res) => {

  const results = [];
  const csvFilePath = path.join(__dirname, "public", "project.csv");
  const data = readDataFromFile();
  fs.createReadStream(csvFilePath)
    .pipe(csv())
    .on("data", (data) => results.push(data))
    .on("end", () => {
      // res.json({results,...data});
      res.send({results,...data})
    })
    .on("error", (err) => {
      res.status(500).json({ error: "Failed to read CSV file" });
      console.error(err);
    });

    
});

app.post("/api/submit", (req, res) => {
  const { projectScores, groupId, studentId, email, groupDescription } = req.body;
  const excelFilePath = './public/total.xlsx';

  let workbook;
  if (fs.existsSync(excelFilePath)) {
    workbook = XLSX.readFile(excelFilePath);
  } else {
    workbook = XLSX.utils.book_new();
  }

  const worksheet = workbook.Sheets['Sheet1'] || XLSX.utils.aoa_to_sheet([[]]);

  const existingData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
  const lastRow = existingData.length;

  const newData = [
    [
      groupId,
      studentId,
      email,
      groupDescription,
      ...projectScores.reduce((acc, score) => {
        return acc.concat([score.project, score.score]);
      }, []),
    ],
  ];

  XLSX.utils.sheet_add_aoa(worksheet, newData, { origin: lastRow });

  if (!workbook.Sheets['Sheet1']) {
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  }

  XLSX.writeFile(workbook, excelFilePath);
  
  res.status(200).json({ msg: "Success" });
});


app.get("/api/Allproject", (req, res) => {
  const excelFilePath = path.join(__dirname, "public", "total.xlsx");
  const data2 = readDataFromFile();
  try {
    const workbook = XLSX.readFile(excelFilePath);

    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    res.json({data:data,data2:data2});
  } catch (err) {
    res.status(500).json({ error: "Failed to read total file" });
    console.error(err);
  }
});
app.get('/api/download', (req, res) => {
  const filePath = path.join(__dirname, 'public', 'total.xlsx');
  
  res.download(filePath, 'total.xlsx', (err) => {
    if (err) {
      console.error('Error downloading file:', err);
      res.status(500).json({ error: 'Failed to download the file' });
    }
  });
});
app.listen(port, '0.0.0.0', () => {
  console.log(`App is available on http://0.0.0.0:${port}`);
});
