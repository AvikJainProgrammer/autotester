const xlsx = require('xlsx');
const fetch = require('node-fetch');

async function sendQuestionAndGetResponse(question) {
  const apiUrl = 'https://azdogrcogsearchpunjabidev.azurewebsites.net/get_response'; // Update with the correct API URL
  try {
    const response = await fetch(apiUrl, {
      method: "POST",
      body: JSON.stringify({ message: question }),
      headers: { "Content-Type": "application/json" },
    });
    const data = await response.json();
    console.log(data.assistant_content);
    return data.assistant_content; // Adjust this depending on the response structure
  } catch (error) {
    console.error("Error:", error);
    return null;
  }
}

async function processExcel() {
  const workbook = xlsx.readFile('questions.xlsx'); // Replace with your input Excel file name
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const questions = xlsx.utils.sheet_to_json(worksheet, { header: 1 }).slice(1).map(row => row[0]);

  const results = [];

  for (let question of questions) {
    const answer = await sendQuestionAndGetResponse(question);
    results.push({ Question: question, Answer: answer });
  }

  const newWorkbook = xlsx.utils.book_new();
  const newWorksheet = xlsx.utils.json_to_sheet(results);
  xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Results');
  xlsx.writeFile(newWorkbook, 'output.xlsx'); // Replace with your desired output file name
}

processExcel();

