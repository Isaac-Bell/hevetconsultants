const axios = require('axios');
const fs = require('fs');
const xlsx = require('xlsx');
const cheerio = require('cheerio');

const workbook = xlsx.readFile('builders.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

const updatedRows = [];

async function fetchData(url) {
  try {
    const response = await axios.get(url);
    const $ = cheerio.load(response.data);
    // Extract the email and phone data from the webpage using Cheerio selectors
    const email = $('#olp_primaryemailpr').text();
    const phone = $('#olp_mobilenumberpr').text();
    return { email, phone };
  } catch (error) {
    console.error('Error fetching data:', error);
    return { email: '', phone: '' };
  }
}

async function populateData() {
  for (let i = 1; i < rows.length; i++) {
    const url = rows[i][0]; // Assuming the hyperlinked URLs are in column A
    const emailColumn = `C${i + 1}`; // Assuming the email address column is C
    const phoneColumn = `D${i + 1}`; // Assuming the mobile number column is D

    const { email, phone } = await fetchData(url);

    worksheet[emailColumn] = { t: 's', v: email };
    worksheet[phoneColumn] = { t: 's', v: phone };

    updatedRows.push([url, email, phone]);
  }

  const updatedWorkbook = xlsx.utils.book_new();
  const updatedWorksheet = xlsx.utils.aoa_to_sheet(updatedRows);
  xlsx.utils.book_append_sheet(updatedWorkbook, updatedWorksheet, 'Sheet1');
  xlsx.writeFile(updatedWorkbook, 'output.xlsx');

  console.log('Data population completed.');
}

populateData();
