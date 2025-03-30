function generateWritingQuestions() {
  var sheetName = "study1"; // Change this to your specific sheet name
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (!sheet) {
    Logger.log("Sheet not found: " + sheetName);
    return [];
  }

  var data = sheet.getDataRange().getValues(); // Read all data

  if (data.length < 2) {
    Logger.log("Not enough data in the sheet.");
    return [];
  }

  var wordsAndExpressions = data.slice(1).flatMap(row => row).join(", "); // Extract words & expressions

  var apiKey = "Your API"; // Replace with your actual API key
  var apiUrl = "https://api.x.ai/v1/chat/completions";

  var payload = {
    messages: [
      { role: "system", content: "You are an advanced English writing tutor." },
      {
        role: "user",
        content: `Based on the following words and expressions, generate five questions to help improve English skills. 
                  I saved Korean expressions that I can remember when remembering English expresisons. Use the Korean words 
                  to make questions; use the english sentence to make questions. Ramdomly choose five questions for the day.

                  Words and expressions: ${wordsAndExpressions}

                  Each question should include a specific question that I can answer and proper usage of the words.`
      }
    ],
    model: "grok-2-latest",
    stream: false,
    temperature: 0.7
  };

  var options = {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": `Bearer ${apiKey}` },
    payload: JSON.stringify(payload)
  };

  try {
    var response = UrlFetchApp.fetch(apiUrl, options);
    var json = JSON.parse(response.getContentText());
    var questions = json.choices[0].message.content.split("\n").filter(q => q.trim() !== ""); // Extract questions
    return questions;
  } catch (e) {
    Logger.log("Error fetching questions: " + e.toString());
    return [];
  }
}

function sendWritingEmail() {
  var questions = generateWritingQuestions();
  if (questions.length === 0) {
    Logger.log("No questions generated.");
    return;
  }

  var subject = "Daily English Writing Questions";
  var body = "Here are your writing questions for today:\n\n" + questions.join("\n\n");

  MailApp.sendEmail("jin9465@gmail.com", subject, body);
}

function scheduleDailyWritingEmail() {
  ScriptApp.newTrigger("sendWritingEmail")
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();
}

// Function to generate questions on demand and display them in a sheet
function generateQuestionsOnDemand() {
  var questions = generateWritingQuestions();
  if (questions.length === 0) {
    SpreadsheetApp.getUi().alert("No questions generated.");
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Generated Questions") || ss.insertSheet("Generated Questions");

  sheet.clear(); // Clear previous questions
  sheet.appendRow(["Generated Writing Questions"]); // Header
  questions.forEach(q => sheet.appendRow([q])); // Add questions to sheet

  SpreadsheetApp.getUi().alert("Questions generated and added to 'Generated Questions' sheet.");
}
