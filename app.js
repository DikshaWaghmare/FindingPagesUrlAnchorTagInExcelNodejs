const express = require('express');
const axios = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs'); // Using 'exceljs' for Excel file creation
const { readFileSync } = require('fs');

const app = express();
app.set('view engine', 'ejs');
app.use(express.urlencoded({ extended: true }));

async function extractAnchorTags(urlList) {
  const data = [];
  for (const url of urlList) {
    try {
      const response = await axios.get(url);
      const $ = cheerio.load(response.data);
      $('a').each((_, element) => {
        const link = $(element).attr('href');
        const text = $(element).text();
        data.push({ URL: url, Anchor_Text: text, Link: link });
      });
    } catch (error) {
      console.error(`Error accessing ${url}: ${error}`);
    }
  }
  return data;
}

app.get('/', (req, res) => {
  res.render('index');
});

app.post('/', async (req, res) => {
  const urls = req.body.urls.split('\n');
  const extractedData = await extractAnchorTags(urls);

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Extracted Data');
  worksheet.columns = [
    { header: 'URL', key: 'URL', width: 40 },
    { header: 'Anchor Text', key: 'Anchor_Text', width: 40 },
    { header: 'Link', key: 'Link', width: 40 },
  ];

  extractedData.forEach((row) => {
    worksheet.addRow(row);
  });

  res.setHeader(
    'Content-Type',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );
  res.setHeader('Content-Disposition', 'attachment; filename=extracted_data.xlsx');

  const buffer = await workbook.xlsx.writeBuffer();
  res.send(buffer);
});

const port = process.env.PORT || 5000;
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
