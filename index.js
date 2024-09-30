const express = require('express');
const multer = require('multer');
const app = express();
const cors = require('cors');
const {
  ServicePrincipalCredentials,
  PDFServices,
  MimeType,
  ExtractPDFParams,
  ExtractElementType,
  ExtractPDFJob,
  ExtractPDFResult,
  SDKError,
  ServiceUsageError,
  ServiceApiError
} = require("@adobe/pdfservices-node-sdk");
const fs = require('fs');
const path = require("path");
const AdmZip = require('adm-zip');
const XLSX = require('xlsx');

// In-memory data store
let resultsStore = [];

app.use(cors({
  origin: 'http://127.0.0.1:5500', // Adjust this to your frontend's origin
  methods: ['POST', 'GET', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  credentials: true,
}));

// Configure multer for handling file uploads
const upload = multer({ dest: 'uploads/' });

app.post('/extract-text', upload.array('pdfFiles'), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).send('No files uploaded');
    }

    const credentials = new ServicePrincipalCredentials({
      clientId: 'd8ac4256d50c48f9a66650bee6f95acf',
      clientSecret: 'p8e-KWUxu2lR5gKjSJhLdU8AwQLY4F5tdH0r'
    });

    // Increase the timeout to 30 seconds
    const pdfServices = new PDFServices({ 
      credentials, 
      timeout: 30000  // Increase timeout to 30 seconds
    });

    const results = await Promise.all(req.files.map(async (file) => {
      try {
        const readStream = fs.createReadStream(file.path);

        const inputAsset = await pdfServices.upload({
          readStream,
          mimeType: MimeType.PDF
        });

        const params = new ExtractPDFParams({
          elementsToExtract: [ExtractElementType.TEXT]
        });

        const job = new ExtractPDFJob({ inputAsset, params });

        const pollingURL = await pdfServices.submit({ job });
        const pdfServicesResponse = await pdfServices.getJobResult({
          pollingURL,
          resultType: ExtractPDFResult
        });

        const resultAsset = pdfServicesResponse.result.resource;
        const streamAsset = await pdfServices.getContent({ asset: resultAsset });

        const outputFilePath = createOutputFilePath();

        // Ensure the file is fully written before proceeding
        await new Promise((resolve, reject) => {
          const writeStream = fs.createWriteStream(outputFilePath);
          streamAsset.readStream.pipe(writeStream);
          writeStream.on('finish', resolve);
          writeStream.on('error', reject);
        });

        const zip = new AdmZip(outputFilePath);
        const zipEntries = zip.getEntries();
        const jsonFile = zipEntries.find(entry => entry.entryName.endsWith('.json'));

        if (!jsonFile) {
          throw new Error('No JSON file found in the zip');
        }

        const jsonContent = JSON.parse(zip.readAsText(jsonFile));

        const extractedTexts = jsonContent.elements
          .filter(element => element.Text)
          .map(element => element.Text);

        await fs.promises.unlink(outputFilePath);
        await fs.promises.unlink(file.path);  // Clean up the uploaded file

        const result = processExtractedTexts(extractedTexts);
        resultsStore.push(result); // Store the result
        return result;
      } catch (err) {
        console.error(err);
        throw err;
      }
    }));
    console.log(results);
    resultsStore = results;
    saveToExcel(results);
    res.json(results);
  } catch (err) {
    console.log("Exception encountered while executing operation", err);
    res.status(500).json({ error: `Error extracting text: ${err.message} `});
  }
});

// GET endpoint to retrieve stored results
app.get('/results', (req, res) => {
  res.json(resultsStore);
});

function processExtractedTexts(texts) {
  let result = {};

  texts.forEach(text => {
    if (text.includes('Data Analytics with Python')) {
      result.course = 'Data Analytics with Python';
      const matches = text.match(/(\w+\s\w+)\s([\d.]+\/\d+)\s([\d.]+\/\d+)\s(\d+)/);
      if (matches) {
        result.name = matches[1];
        result.score = matches[4];
      }
    } else if (text.startsWith('Roll No:')) {
      result.rollNo = text.split(':')[1].trim();
    } else if (text.includes('credits recommended')) {
      result.credits = text.split(':')[1].trim();
    }
  });

  return {
    heading: result.course,
    name: result.name,
    score: result.score,
    rollNo: result.rollNo,
    credits: result.credits
  };
}

function createOutputFilePath() {
  const filePath = "output/ExtractTextInfoFromPDF/";
  const date = new Date();
  const dateString = date.getFullYear() + "-" + ("0" + (date.getMonth() + 1)).slice(-2) + "-" +
    ("0" + date.getDate()).slice(-2) + "T" + ("0" + date.getHours()).slice(-2) + "-" +
    ("0" + date.getMinutes()).slice(-2) + "-" + ("0" + date.getSeconds()).slice(-2);
  
  try {
    fs.mkdirSync(filePath, { recursive: true });  // Create the directory if it doesn't exist
  } catch (err) {
    if (err.code !== 'EEXIST') throw err;  // Only throw if the error is not because the directory already exists
  }

  return (`${filePath}extract${dateString}.zip`);
}
function saveToExcel(results) {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(results);
  
  XLSX.utils.book_append_sheet(workbook, worksheet, "Extraction Results");
  
  const excelFilePath = path.join(__dirname, 'output', 'extraction_results.xlsx');
  
  // Ensure the output directory exists
  const outputDir = path.dirname(excelFilePath);
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }
  
  XLSX.writeFile(workbook, excelFilePath);
  console.log(`Results saved to ${excelFilePath}`);
}

app.listen(3000, () => {
  console.log('Server started on port 3000');
});