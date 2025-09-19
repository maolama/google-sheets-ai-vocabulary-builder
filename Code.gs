/**
 * This Google Apps Script automates vocabulary building by fetching
 * explanations, examples, and translations for words entered into a Google Sheet
 * directly from the Gemini API. It also includes functions for visual formatting and quizzes.
 */

// --- Configuration ---
const SPREADSHEET_ID = 'YOUR_SPREED_SHEET_ID';
const SHEET_NAME = 'SHEET_NAME';
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const GEMINI_MODEL = 'gemini-2.5-flash';

function setApiKeys() {
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', 'YOUR_GEMINI_API_KEY');
  Logger.log('Gemini API Key set successfully in script properties.');
}

// --- Formatting Constants ---
const EVEN_ROW_COLOR = '#f3f3f3';
const ODD_ROW_COLOR = '#ffffff';
const PAST_DUE_COLOR = '#ffcdd2';
const DUE_TODAY_COLOR = '#ffcc80';
const DUE_TOMORROW_COLOR = '#fff9c4';
const BROKEN_LINK_COLOR = '#ffcdd2'; // Light Red for broken links


// The number of words to process in each batch. Two concurrent API calls per batch.
const BATCH_SIZE = 9; 
const SCRIPT_PROPERTY_KEY = 'LAST_PROCESSED_ROW';


/**
 * Triggered automatically when a cell in the spreadsheet is edited.
 */
function onEdit(e) {
  const lock = LockService.getScriptLock();
  // if (lock.tryLock(10000)) {
    try {
      const range = e.range;
      const sheet = range.getSheet();

      if (sheet.getName() === SHEET_NAME && range.getColumn() === 1 && range.getNumRows() === 1) {
        const word = range.getValue().toString().trim();
        const row = range.getRow();

        if (!word) {
          sheet.getRange(row, 2, 1, 24).clearContent(); // Clear 19 columns (B to T)
          formatRow(sheet, row);
          return;
        }

        // **CHANGE**: Call the shared processing function
        processNewWord(word, row);
      }
    } finally {
      lock.releaseLock();
    }
  // } else {
  //   Logger.log('Could not acquire lock for onEdit trigger.');
  // }
}

/**
 * **NEW**: This is the core logic for processing a new word.
 * It's called by both onEdit and addNewWord.
 * @param {string} word The word to process.
 * @param {number} row The row number to populate.
 */
function processNewWord(word, row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  
  // Duplicate check
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const wordListRange = sheet.getRange(2, 1, lastRow - 1, 1);
    const wordValues = wordListRange.getValues();
    for (let i = 0; i < wordValues.length; i++) {
      if (wordValues[i][0].toString().trim().toLowerCase() === word.toLowerCase() && (i + 2) !== row) {
        SpreadsheetApp.getUi().alert(`The word "${word}" already exists in row ${i + 2}.`);
        sheet.getRange(row, 1).clearContent();
        return;
      }
    }
  }

  // getGeminiDefinitionAndExamples now returns an array of objects
  const geminiResponseArray = getGeminiDefinitionAndExamples(word); 
  const now = new Date();
  
  let allPartsOfSpeech = '', allPersianTranslations = '', allDefinitions = '', allDefinitionExamples = '';
  let allGeneralExamplesEN = '', allGeneralExamplesPER = '', allSynonyms = '', allAntonyms = '';
  let allNotes = '', allWordFamily = '', allUkPronunciation = '', allUsPronunciation = '';
  
  geminiResponseArray.forEach((posData, index) => {
    if (index > 0) {
      const separator = '\n\n';
      allPartsOfSpeech += ', ';
      allPersianTranslations += separator;
      allDefinitions += separator;
      allDefinitionExamples += separator;
      allGeneralExamplesEN += separator;
      allGeneralExamplesPER += separator;
      allSynonyms += separator;
      allAntonyms += separator;
      allNotes += separator;
      allWordFamily += separator;
    }

    const posHeader = ``;
    
    allPartsOfSpeech += posData.partOfSpeech || '—';
    allDefinitions += `${posHeader}\n` + (posData.meanings || []).map(m => `• ${m.definition}`).join('\n');
    allPersianTranslations += `${posHeader}\n` + (posData.meanings || []).map(m => `• ${m.persianTranslation}`).join('\n');
    allDefinitionExamples += `${posHeader}\n` + (posData.meanings || []).map(m => `• ${m.example}`).join('\n');
    allGeneralExamplesEN += `${posHeader}\n` + (posData.generalExamples || []).map(item => `• ${item.example}`).join('\n');
    allGeneralExamplesPER += `${posHeader}\n` + (posData.generalExamples || []).map(item => `• ${item.translation}`).join('\n');
    allSynonyms += `${posHeader}\n` + (posData.synonyms || []).map(s => `• ${s}`).join('\n');
    allAntonyms += `${posHeader}\n` + (posData.antonyms || []).map(a => `• ${a}`).join('\n');
    allNotes += `${posHeader}\n` + (posData.notes || []).map(n => `• ${n}`).join('\n');
    allWordFamily += `${posHeader}\n` + (posData.wordFamily || []).map(item => `• ${item.word} (${item.POS})`).join('\n');
    
    allUkPronunciation += `${posData.partOfSpeech}:${posData.ukPronunciation || '—'}\n`;
    allUsPronunciation += `${posData.partOfSpeech}:${posData.usPronunciation || '—'}\n`;
  });

  sheet.getRange(row, 2).setValue(allPersianTranslations.trim());
  sheet.getRange(row, 3).setValue(allDefinitions.trim());
  sheet.getRange(row, 4).setValue(allDefinitionExamples.trim());
  sheet.getRange(row, 5).setValue(allGeneralExamplesEN.trim());
  sheet.getRange(row, 6).setValue(allGeneralExamplesPER.trim());
  sheet.getRange(row, 7).setValue(allPartsOfSpeech.trim());
  sheet.getRange(row, 8).setValue(allSynonyms.trim());
  sheet.getRange(row, 9).setValue(allAntonyms.trim());
  sheet.getRange(row, 10).setValue(allNotes.trim());
  sheet.getRange(row, 11).setValue(allWordFamily.trim());
  // Set values for columns L through U (12 to 21)
  const addedDateCell = sheet.getRange(row, 12); // Column L: Created At
  if (!addedDateCell.getValue()) {
    addedDateCell.setValue(now);
  }
  sheet.getRange(row, 13).setValue(now); // Column M: Modified At
  sheet.getRange(row, 14).setValue(allUkPronunciation.trim()); // Column N: UK Pronunciation
  sheet.getRange(row, 15).setValue(allUsPronunciation.trim()); // Column O: US Pronunciation

  const encodedWord = encodeURIComponent(word.toLowerCase().replace(/ /g, '-'));
  sheet.getRange(row, 16).setFormula(`=HYPERLINK("https://dictionary.cambridge.org/dictionary/english/${encodedWord}", "Cambridge")`);
  sheet.getRange(row, 17).setFormula(`=HYPERLINK("https://www.oxfordlearnersdictionaries.com/definition/english/${encodedWord}", "Oxford")`);
  
  const nextReviewDate = new Date(now);
  nextReviewDate.setDate(now.getDate() + 1);
  sheet.getRange(row, 18).setValue(nextReviewDate);   // Column R: Next Review
  sheet.getRange(row, 19).setValue(0);              // Column S: Review Count
  sheet.getRange(row, 20).setValue(0);              // Column T: Quiz Count
  sheet.getRange(row, 21).setValue(0);              // Column U: Total Reviews

  // Initialize new metadata columns V, W, X (22, 23, 24)
  sheet.getRange(row, 22).setValue(0); // Column V: Writing
  sheet.getRange(row, 23).setValue(0); // Column W: Speaking
  sheet.getRange(row, 24).setValue(0); // Column X: Difficulty
  
  formatRow(sheet, row);
}

function getDefaultGeminiResponse(errorMsg = 'Error') {
  // This function returns a single object that represents a default/error state.
  // The calling function, getGeminiDefinitionAndExamples, will wrap this in an array.
  return {
    partOfSpeech: errorMsg,
    meanings: [{ definition: '—', example: '—', persianTranslation: '—' }],
    generalExamples: [{ example: '—', translation: '—' }],
    synonyms: ['—'], // Must be an array
    antonyms: ['—'], // Must be an array
    notes: ['—'],    // Must be an array
    ukPronunciation: '—',
    usPronunciation: '—',
    wordFamily: [{ word: '—', POS: '—' }]
  };
}


/**
 * **NEW**: Prepares the request object for a Gemini API call for a single word.
 * This does NOT execute the request.
 * @param {string} word The word to get data for.
 * @returns {Object} A URL Fetch request object.
 */
function prepareGeminiRequest(word) {
  // This is the prompt from your getGeminiDefinitionAndExamples function
  const prompt = `Analyze the English word "${word}". Your response MUST be a single, valid JSON ARRAY and nothing else.
Each object in the array should represent a single part of speech for the word.
For each part of speech, provide a complete analysis. Do not include any explanatory text, markdown, or comments outside the JSON.

The structure for EACH object in the array must be:
{
  "partOfSpeech": "string",
  "meanings": [
    {
      "definition": "string",        // a plain, simple explanation
      "example": "string",           // a single, clear example sentence for that specific definition
      "persianTranslation": "string" // the short Persian word equivalent for that specific meaning
    }
  ],
  "generalExamples": [ // An array of objects. Provide at least 10 examples. For the "example" key, provide a diverse set of sentences. A diverse set includes: a mix of simple and complex(compund) sentences; for verbs, different tenses (past, present, future) and voices (active, passive); for nouns, singular and plural forms. Keep the same dicersity for adjectives and adverbs. The "translation" key is its FLUENT Persian translation.
    {
      "example": "string",
      "translation": "string"
    }
  ],
  "synonyms": ["string"],      // An array of up to 5 "common" and "close" synonyms. If none, return an empty array [].
  "antonyms": ["string"],      // An array of up to 5 "common" and "close" antonyms. If none, return an empty array [].
  "notes": ["string"],         // An array of detailed notes. Include common collocations, usage mistakes, context (formal/informal), and helpful tips. If none, return an empty array [].
  "ukPronunciation": "string", // common Articulatory Phonetics spelling for UK English using IPA.
  "usPronunciation": "string", // common Articulatory Phonetics spelling for US English using IPA.
  "wordFamily": [ // An array of objects. Each object must have a "word" key and its corresponding part of speech in a "POS" key. If none, return an empty array [].
    {
      "word": "string",
      "POS": "string"
    }
  ]
}

// --- START OF EXAMPLE ---
// Here is a perfect example for the word "record", which has different noun and verb forms.
[
  {
    "partOfSpeech": "noun",
    "meanings": [
      {
        "definition": "A piece of information or an event that is written down or stored on a computer so it can be looked at in the future.",
        "example": "The school keeps a record of all its students.",
        "persianTranslation": "سابقه"
      },
      {
        "definition": "The best result or performance ever achieved in a particular sport or activity.",
        "example": "She broke the world record for the 100 meters.",
        "persianTranslation": "رکورد"
      }
    ],
    "generalExamples": [
      {
        "example": "According to official records, he was born in 1985.",
        "translation": "طبق سوابق رسمی، او در سال ۱۹۸۵ متولد شده است."
      },
      {
        "example": "The athlete has an impressive track record of wins.",
        "translation": "این ورزشکار سابقه چشمگیری از پیروزی‌ها دارد."
      }, ...
    ],
    "synonyms": ["document", "file", "account", "best performance"],
    "antonyms": [],
    "notes": ["As a noun, the stress is on the first syllable: RE-cord." , ...],
    "ukPronunciation": "/ˈrek.ɔːd/",
    "usPronunciation": "/ˈrek.ɚd/",
    "wordFamily": [{"word": "recorder", "POS": "noun"}]
  },
  {
    "partOfSpeech": "verb",
    "meanings": [
      {
        "definition": "To store sounds, images, or information electronically so that it can be heard or seen later.",
        "example": "We need to record the meeting for those who can't attend.",
        "persianTranslation": "ضبط کردن"
      }
    ],
    "generalExamples": [
      {
        "example": "The band is recording a new album next month.",
        "translation": "گروه ماه آینده آلبوم جدیدی را ضبط خواهد کرد."
      },
      {
        "example": "All conversations on this line are recorded for quality assurance.",
        "translation": "تمام مکالمات این خط برای تضمین کیفیت ضبط می‌شود."
      },
      {
        "example": "He carefully recorded all the expenses in his notebook.",
        "translation": "او با دقت تمام هزینه‌ها را در دفترچه‌اش ثبت کرد."
      }, ...
    ],
    "synonyms": ["document", "log", "register", "tape"],
    "antonyms": ["delete", "erase"],
    "notes": ["As a verb, the stress is on the second syllable: re-CORD.", ...],
    "ukPronunciation": "/rɪˈkɔːd/",
    "usPronunciation": "/rɪˈkɔːrd/",
    "wordFamily": [{"word": "recording", "POS": "noun"}]
  }
]
// --- END OF EXAMPLE ---

Now, provide the JSON array for the word "${word}".`;
  
   const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generation_config: { temperature: 0.7, top_p: 0.95, top_k: 40 }
  };

  return {
    url: `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`,
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': GEMINI_API_KEY },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
}

/**
 * **NEW**: Parses a successful Gemini API response text.
 * @param {string} responseText The raw text content from the API response.
 * @returns {Object} The parsed and structured word data object.
 */
function parseGeminiResponse(responseText) {
  const data = JSON.parse(responseText);
  if (data.candidates && data.candidates[0].content && data.candidates[0].content.parts) {
    const geminiOutputText = data.candidates[0].content.parts[0].text.trim().replace(/```(?:json)?\s*/gi, "").replace(/```$/, "").trim();
    return JSON.parse(geminiOutputText);
  }
  throw new Error("Invalid Gemini response format.");
}



/**
 * **MODIFIED**: Calls the Gemini API for a single word using the new helper functions.
 */
function getGeminiDefinitionAndExamples(word) {
  try {
    const request = prepareGeminiRequest(word);
    const response = UrlFetchApp.fetch(request.url, request); // UrlFetchApp.fetch now takes the request object directly
    const responseCode = response.getResponseCode();

    if (responseCode >= 200 && responseCode < 300) {
      // The AI now returns an array, which we pass directly.
      return parseGeminiResponse(response.getContentText());
    } else {
      // getDefaultGeminiResponse needs to return an array with a default error object
      return [getDefaultGeminiResponse(`API error (HTTP ${responseCode}).`)];
    }
  } catch (e) {
    console.error(`Error during API call for "${word}": ${e.message}`);
    return [getDefaultGeminiResponse('Script or network error.')];
  }
}

function getWordsDueForReview() {
  Logger.log("[Backend] getWordsDueForReview: Starting to fetch words.");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log("[Backend] getWordsDueForReview: Sheet not found.");
    throw new Error(`Sheet with name "${SHEET_NAME}" not found.`);
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("[Backend] getWordsDueForReview: No words in the sheet.");
    return [];
  }

  // **MODIFIED**: Read all 24 columns of data (A-U) instead of 20.
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 24);
  const values = dataRange.getValues();
  // **MODIFIED**: Hyperlinks are now in Columns P and Q (index 16).
  const linkFormulas = sheet.getRange(2, 16, lastRow - 1, 2).getFormulas();

  const wordsDue = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  values.forEach((row, index) => {
    const word = row[0];
    // **MODIFIED**: Next Review Date is now in Column R (index 17).
    const nextReviewDateCell = row[17]; 

    if (word && nextReviewDateCell instanceof Date) {
      const nextReviewDate = new Date(nextReviewDateCell);
      nextReviewDate.setHours(0, 0, 0, 0);

      if (nextReviewDate <= today) {
        const cambridgeFormula = linkFormulas[index][0];
        const oxfordFormula = linkFormulas[index][1];
        
        const cambridgeUrl = cambridgeFormula.match(/HYPERLINK\("([^"]+)"/i)?.[1] || '';
        const oxfordUrl = oxfordFormula.match(/HYPERLINK\("([^"]+)"/i)?.[1] || '';

        wordsDue.push({
          word: word,
          row: index + 2,
          persianTranslations: row[1] || '',
          definitions: row[2] || '',
          definitionExamples: row[3] || '',
          // **MODIFIED**: English examples are in row[4], and new Persian examples in row[5].
          generalExamples: row[4] || '',
          generalExamplesPersian: row[5] || '',
          // **MODIFIED**: All subsequent column indexes are shifted by +1.
          partOfSpeech: row[6],
          synonyms: row[7],
          antonyms: row[8],
          notes: row[9],
          wordFamily: row[10],
          ukPronunciation: row[13],
          usPronunciation: row[14],
          cambridgeUrl: cambridgeUrl,
          oxfordUrl: oxfordUrl,
          reviewCount: row[18] || 0, // Was 17
          totalReviews: row[20] || 0,  // Was 19
          // **NEW**: Add the new properties from the end columns.
          writing: row[21] || 0,    // Column V
          speaking: row[22] || 0,   // Column W
          difficulty: row[23] || 0  // Column X
        });
      }
    }
  });

  Logger.log(`[Backend] getWordsDueForReview: Found ${wordsDue.length} words due for review today.`);
  return wordsDue;
}

/**
 * Implements new review count logic.
 */
function updateWordReview(update) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const now = new Date();
    const { row, difficulty, reviewCount, totalReviews } = update;
    
    let currentReviewCount = Number(reviewCount) || 0;
    let currentTotalReviews = Number(totalReviews) || 0;
    
    let newReviewCount = currentReviewCount;
    let newTotalReviews = currentTotalReviews;
    let daysToAdd = 0;

    if (typeof difficulty === 'number' && difficulty > 0) {
      daysToAdd = difficulty;
      newReviewCount = currentReviewCount + 1;
      newTotalReviews = currentTotalReviews + 1;
    } 
    else if (typeof difficulty === 'string') {
      switch (difficulty) {
        case 'Again': 
          daysToAdd = 0;
          newReviewCount = 0;
          break;
        case 'Hard': 
          daysToAdd = 1;
          newTotalReviews = currentTotalReviews + 1;
          break;
        case 'Good': 
          daysToAdd = (currentTotalReviews < 3) ? [3, 7, 14][currentTotalReviews] : 30; 
          newReviewCount++; 
          newTotalReviews++;
          break;
        case 'Easy': 
          daysToAdd = (currentTotalReviews < 3) ? [7, 30, 90][currentTotalReviews] : 180; 
          newReviewCount += 2; 
          newTotalReviews++;
          break;
        default: console.warn(`Unknown difficulty string: ${difficulty}.`); return;
      }
    } else {
      console.warn(`Invalid difficulty type: ${difficulty}.`);
      return;
    }

    const newNextReviewDate = new Date();
    newNextReviewDate.setDate(newNextReviewDate.getDate() + daysToAdd);
    newNextReviewDate.setHours(0, 0, 0, 0);

    // **MODIFIED**: Updated column indexes to reflect the new layout.
    sheet.getRange(row, 13).setValue(now);                      // Modified At: Was 12
    sheet.getRange(row, 18).setValue(newNextReviewDate);      // Next Review Time: Was 17
    sheet.getRange(row, 19).setValue(newReviewCount);         // Review Count: Was 18
    sheet.getRange(row, 21).setValue(newTotalReviews);        // Total Reviews: Was 20
    
    formatRow(sheet, row);
    
    console.log(`Asynchronously updated row ${row} for difficulty '${difficulty}'.`);
  } catch (e) {
    console.error(`Failed to update row ${update ? update.row : 'unknown'}. Error: ${e.message}`);
  }
}

/**
 * Formats a single row with alternating colors and conditional formatting.
 */
function formatRow(sheet, rowNumber) {
  // **MODIFIED**: The range now covers all 24 columns.
  const wholeRowRange = sheet.getRange(rowNumber, 1, 1, 24);
  const wordCell = sheet.getRange(rowNumber, 1);
  // **MODIFIED**: Review Date is now in Column R (18).
  const reviewDateCell = sheet.getRange(rowNumber, 18);

  if (wordCell.getValue() === '') {
    wholeRowRange.setBackground(ODD_ROW_COLOR);
    return;
  }

  const rowBgColor = (rowNumber % 2 === 0) ? EVEN_ROW_COLOR : ODD_ROW_COLOR;
  wholeRowRange.setBackground(rowBgColor);

  const reviewDateVal = reviewDateCell.getValue();
  if (reviewDateVal instanceof Date) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const tomorrow = new Date(today);
    tomorrow.setDate(today.getDate() + 1);

    const reviewDate = new Date(reviewDateVal);
    reviewDate.setHours(0, 0, 0, 0);

    if (reviewDate < today) {
      wordCell.setBackground(PAST_DUE_COLOR);
    } else if (reviewDate.getTime() === today.getTime()) {
      wordCell.setBackground(DUE_TODAY_COLOR);
    } else if (reviewDate.getTime() === tomorrow.getTime()) {
      wordCell.setBackground(DUE_TOMORROW_COLOR);
    } else {
      wordCell.setBackground(rowBgColor);
    }
  }
}

/**
 * Sorts the sheet using the built-in, safer sort method and then applies formatting.
 */
function formatSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // **MODIFIED**: The range now covers all 24 columns.
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 24);
  
  // **MODIFIED**: Sort columns have been updated.
  dataRange.sort([
    { column: 18, ascending: true }, // Sort by Review Date (Column R, was 17)
    { column: 19, ascending: true }, // Then by Review Count (Column S, was 18)
    { column: 1, ascending: true }   // Then by Word (Column A)
  ]);
  
  SpreadsheetApp.flush();

  // Now that it's sorted, read the values again to apply colors
  const values = dataRange.getValues();
  const backgroundColors = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  for (let i = 0; i < values.length; i++) {
    const rowNumber = i + 2;
    const rowData = values[i];
    const word = rowData[0];
    // **MODIFIED**: Review Date is at index 17 in the values array (Column R).
    const reviewDateVal = rowData[17];

    const rowBgColor = (rowNumber % 2 === 0) ? EVEN_ROW_COLOR : ODD_ROW_COLOR;
    // **MODIFIED**: Color arrays are now 24 elements long.
    const rowColors = new Array(24).fill(rowBgColor);

    if (word === '') {
      backgroundColors.push(new Array(24).fill(ODD_ROW_COLOR));
      continue;
    }

    if (reviewDateVal instanceof Date) {
      const reviewDate = new Date(reviewDateVal);
      reviewDate.setHours(0, 0, 0, 0);

      if (reviewDate < today) {
        rowColors[0] = PAST_DUE_COLOR;
      } else if (reviewDate.getTime() === today.getTime()) {
        rowColors[0] = DUE_TODAY_COLOR;
      } else if (reviewDate.getTime() === tomorrow.getTime()) {
        rowColors[0] = DUE_TOMORROW_COLOR;
      }
    }
    backgroundColors.push(rowColors);
  }

  if (backgroundColors.length > 0) {
    dataRange.setBackgrounds(backgroundColors);
  }
  
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Sheet has been sorted and formatted!');
}

/**
 * Sorts the sheet using the built-in, safer sort method and then applies formatting.
 */
function formatSheet_onlyColor() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // **MODIFIED**: The range now covers all 24 columns.
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 24);
  
  // Now that it's sorted, read the values again to apply colors
  const values = dataRange.getValues();
  const backgroundColors = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  for (let i = 0; i < values.length; i++) {
    const rowNumber = i + 2;
    const rowData = values[i];
    const word = rowData[0];
    // **MODIFIED**: Review Date is at index 17 in the values array (Column R).
    const reviewDateVal = rowData[17];

    const rowBgColor = (rowNumber % 2 === 0) ? EVEN_ROW_COLOR : ODD_ROW_COLOR;
    // **MODIFIED**: Color arrays are now 24 elements long.
    const rowColors = new Array(24).fill(rowBgColor);

    if (word === '') {
      backgroundColors.push(new Array(24).fill(ODD_ROW_COLOR));
      continue;
    }

    if (reviewDateVal instanceof Date) {
      const reviewDate = new Date(reviewDateVal);
      reviewDate.setHours(0, 0, 0, 0);

      if (reviewDate < today) {
        rowColors[0] = PAST_DUE_COLOR;
      } else if (reviewDate.getTime() === today.getTime()) {
        rowColors[0] = DUE_TODAY_COLOR;
      } else if (reviewDate.getTime() === tomorrow.getTime()) {
        rowColors[0] = DUE_TOMORROW_COLOR;
      }
    }
    backgroundColors.push(rowColors);
  }

  if (backgroundColors.length > 0) {
    dataRange.setBackgrounds(backgroundColors);
  }
  
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Sheet has been sorted and formatted!');
}



/**
 * Finds and removes duplicate words, keeping the first occurrence.
 */
function removeDuplicateWords() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const dataRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const words = dataRange.getValues();
  
  const seenWords = {};
  const rowsToDelete = [];

  words.forEach((row, index) => {
    const word = row[0].toString().trim().toLowerCase();
    const rowNumber = index + 2;
    if (word) {
      if (seenWords[word]) {
        rowsToDelete.push(rowNumber);
      } else {
        seenWords[word] = true;
      }
    }
  });

  if (rowsToDelete.length > 0) {
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(rowsToDelete[i]);
    }
    SpreadsheetApp.getUi().alert(`${rowsToDelete.length} duplicate word(s) found and removed.`);
    formatSheet();
  } else {
    SpreadsheetApp.getUi().alert('No duplicate words found.');
  }
}

// --- QUIZ FUNCTIONS ---

/**
 * Gets a list of words for the quiz, ensuring they have been reviewed at least once.
 */
function getQuizWords(numQuestions) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  
  // **MODIFIED**: The range now covers all 24 columns.
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 24);
  const allWordsData = dataRange.getValues();

  let allWords = allWordsData.map((row, index) => ({
    word: row[0],
    definition: row[2] ? row[2].split('\n')[0].replace(/• /g, '') : 'No definition',
    // **MODIFIED**: Updated column indexes.
    reviewCount: row[18] || 0, // Column S (was 17)
    quizCount: row[19] || 0,   // Column T (was 18)
    row: index + 2
  })).filter(w => w.word && w.definition && w.reviewCount > 0);

  if (allWords.length < 4) {
    throw new Error("You need at least 4 words that have been reviewed once to start a quiz.");
  }
  
  const minQuizCount = Math.min(...allWords.map(w => w.quizCount));
  const unquizzedWords = allWords.filter(w => w.quizCount === minQuizCount);

  let wordsToQuizFrom = unquizzedWords;
  if (unquizzedWords.length < numQuestions) {
    wordsToQuizFrom = allWords;
  }
  
  const shuffled = wordsToQuizFrom.sort(() => 0.5 - Math.random());
  const selectedWords = shuffled.slice(0, numQuestions);

  const questions = selectedWords.map(correctWord => {
    const distractors = allWords
      .filter(w => w.word !== correctWord.word)
      .sort(() => 0.5 - Math.random())
      .slice(0, 3);

    const options = [correctWord, ...distractors].map(w => w.word);
    
    return {
      question: correctWord.definition,
      options: options.sort(() => 0.5 - Math.random()),
      answer: correctWord.word,
      row: correctWord.row
    };
  });

  return questions;
}

/**
 * Updates the sheet after a quiz is completed.
 */
function updateQuizResults(quizResults) {
  const { incorrectRows, allQuizRows } = quizResults;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const now = new Date();
  const tomorrow = new Date();
  tomorrow.setDate(now.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0);

  if (incorrectRows && incorrectRows.length > 0) {
    incorrectRows.forEach(rowNum => {
      // **MODIFIED**: Updated column indexes.
      sheet.getRange(rowNum, 18).setValue(tomorrow); // Next Review Time: Column R (was 17)
      sheet.getRange(rowNum, 13).setValue(now);      // Modified At: Column M (was 12)
      formatRow(sheet, rowNum);
    });
  }

  if (allQuizRows && allQuizRows.length > 0) {
    allQuizRows.forEach(rowNum => {
      // **MODIFIED**: Updated column index for Quiz Count.
      const countCell = sheet.getRange(rowNum, 20); // Column T (was 19)
      const currentCount = countCell.getValue() || 0;
      countCell.setValue(currentCount + 1);
    });
  }
  
  const lastRow = sheet.getLastRow();
  // **MODIFIED**: Updated range for reading and resetting quiz counts.
  const quizCountsRange = sheet.getRange(2, 20, lastRow - 1, 1); // Column T (was 19)
  const quizCounts = quizCountsRange.getValues().flat();
  if (quizCounts.every(count => count > 0)) {
    quizCountsRange.setValue(0);
    Logger.log("All words have been quizzed. Resetting quiz counts.");
  }
}

// --- CRAM MODE FUNCTION ---
/**
 * Gets a list of words scheduled for future review.
 * @param {number} count The number of future words to fetch.
 * @returns {Array<Object>} An array of word objects for review.
 */
function getFutureWords(count) {
  try {
    Logger.log('[Backend] getFutureWords: Starting to fetch ' + count + ' words.');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet with name "${SHEET_NAME}" not found.`);
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    // **MODIFIED**: The range now covers all 24 columns.
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 24);
    const values = dataRange.getValues();
    // **MODIFIED**: Hyperlinks are now in Columns P and Q.
    const linkFormulas = sheet.getRange(2, 16, lastRow - 1, 2).getFormulas();

    const futureWords = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    values.forEach((row, index) => {
      const word = row[0];
      // **MODIFIED**: Next Review Date is now in Column R (index 17).
      const nextReviewDateCell = row[17];

      if (word && nextReviewDateCell instanceof Date) {
        const nextReviewDate = new Date(nextReviewDateCell);
        nextReviewDate.setHours(0, 0, 0, 0);

        if (nextReviewDate > today) {
          const cambridgeFormula = linkFormulas[index][0];
          const oxfordFormula = linkFormulas[index][1];
          
          const cambridgeUrl = cambridgeFormula.match(/HYPERLINK\("([^"]+)"/i)?.[1] || '';
          const oxfordUrl = oxfordFormula.match(/HYPERLINK\("([^"]+)"/i)?.[1] || '';

          // **MODIFIED**: All column indexes updated, and new Persian examples added.
          futureWords.push({
            word: word,
            row: index + 2,
            persianTranslations: row[1] || '',
            definitions: row[2] || '',
            definitionExamples: row[3] || '',
            generalExamples: row[4] || '',
            generalExamplesPersian: row[5] || '',
            partOfSpeech: row[6],
            synonyms: row[7],
            antonyms: row[8],
            notes: row[9],
            wordFamily: row[10],
            ukPronunciation: row[13],
            usPronunciation: row[14],
            cambridgeUrl: cambridgeUrl,
            oxfordUrl: oxfordUrl,
            reviewCount: row[18] || 0,
            totalReviews: row[20] || 0,
            nextReviewDate: nextReviewDate // Keep as Date object for sorting
          });
        }
      }
    });

    Logger.log('[Backend] getFutureWords: Found ' + futureWords.length + ' total future words.');
    futureWords.sort((a, b) => a.nextReviewDate.getTime() - b.nextReviewDate.getTime());
    
    const wordsToReturn = futureWords.slice(0, count).map(word => {
      // Convert date to string AFTER sorting for safe transfer
      word.nextReviewDate = word.nextReviewDate.toISOString();
      return word;
    });

    Logger.log('[Backend] getFutureWords: Returning ' + wordsToReturn.length + ' words for cram mode.');
    return wordsToReturn;
  } catch (e) {
    Logger.log(`[Backend] Error in getFutureWords: ${e.message}`);
    return null; // Return null on error
  }
}


// --- UTILITY FUNCTIONS ---

/**
 * Sets up the sheet with the correct headers.
 */
function initializeSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet with name "${SHEET_NAME}" not found.`);
    return;
  }
  
  // **MODIFIED**: Added the new "Example (Persian)" header and renamed "Example".
  const headers = [
    'Word', 'Persian Translation', 'Definition', 'Definition Example', 'Example (English)',
    'Example (Persian)', 'Part of Speech', 'Synonyms', 'Antonyms', 'Tips', 'Word Family',
    'Created At', 'Modified At', 'UK Pronunciation', 'US Pronunciation',
    'Cambridge', 'Oxford', 'Next Review Time', 'Review Count', 'Quiz Count', 'Total Reviews',
    'Writing', 'Speaking', 'Difficulty'
  ];
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).clearContent();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  SpreadsheetApp.getUi().alert('Sheet has been initialized with the correct headers.');
}

/**
 * Initializes the Quiz Count for all existing words to 0.
 */
function initializeQuizCounts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // **MODIFIED**: Quiz Count is now in Column T (20).
  const quizCountRange = sheet.getRange(2, 20, lastRow - 1, 1);
  const values = quizCountRange.getValues();
  
  let updated = 0;
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === '') {
      values[i][0] = 0;
      updated++;
    }
  }
  
  quizCountRange.setValues(values);
  SpreadsheetApp.getUi().alert(`Initialization complete. ${updated} word(s) were updated with a quiz count of 0.`);
}

/**
 * Initializes the Total Reviews count for all existing words to 0.
 */
function initializeTotalReviews() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // **MODIFIED**: Total Reviews is now in Column U (21).
  const totalReviewsRange = sheet.getRange(2, 21, lastRow - 1, 1);
  const values = totalReviewsRange.getValues();
  
  let updated = 0;
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === '') {
      values[i][0] = 0;
      updated++;
    }
  }
  
  totalReviewsRange.setValues(values);
  SpreadsheetApp.getUi().alert(`Initialization complete. ${updated} word(s) were updated with a total reviews count of 0.`);
}

/**
 * **NEW**: Updates the Writing, Speaking, and Difficulty for a given word.
 */
function updateWordMetadata(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    // **MODIFIED**: Write to the new columns V, W, and X.
    sheet.getRange(data.row, 22).setValue(data.writing);     // Column V: Writing
    sheet.getRange(data.row, 23).setValue(data.speaking);    // Column W: Speaking
    sheet.getRange(data.row, 24).setValue(data.difficulty);  // Column X: Difficulty
    Logger.log(`Updated metadata for row ${data.row}`);
  } catch (e) {
    Logger.log(`Failed to update metadata for row ${data.row}. Error: ${e.message}`);
  }
}



/**
 * Schedules the review dates for all words in batches.
 */
function scheduleInitialReviews() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const wordsPerDay = 5;
  // **MODIFIED**: Next Review Time is now in Column R (18).
  const reviewDateRange = sheet.getRange(2, 18, lastRow - 1, 1);
  const dates = [];
  let dayOffset = 0;

  for (let i = 0; i < lastRow - 1; i++) {
    if (i > 0 && i % wordsPerDay === 0) {
      dayOffset++;
    }
    const reviewDate = new Date();
    reviewDate.setHours(0, 0, 0, 0);
    reviewDate.setDate(reviewDate.getDate() + dayOffset);
    dates.push([reviewDate]);
  }

  reviewDateRange.setValues(dates);
  formatSheet(); // Re-sort and re-color the sheet after scheduling
  SpreadsheetApp.getUi().alert(`Initial review dates have been scheduled for ${lastRow - 1} words.`);
}


/**
 * Checks if a URL is valid by fetching its response code.
 */
function checkLinkValidity(url) {
  return true;
  // try {
  //   const response = UrlFetchApp.fetch(url, {
  //     muteHttpExceptions: true,
  //     method: 'HEAD'
  //   });
  //   return response.getResponseCode() === 200;
  // } catch (e) {
  //   Logger.log(`Error checking link ${url}: ${e.message}`);
  //   return false;
  // }
}

/**
 * Populates and verifies Cambridge and Oxford hyperlinks.
 */
function populateHyperlinks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const wordRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const words = wordRange.getValues();

  const cambridgeFormulas = [];
  const oxfordFormulas = [];
  const cambridgeColors = [];
  const oxfordColors = [];

  words.forEach(row => {
    const word = row[0].toString().trim();
    if (word) {
      const encodedWord = encodeURIComponent(word.toLowerCase().replace(/ /g, '-'));
      const cambridgeUrl = `https://dictionary.cambridge.org/dictionary/english/${encodedWord}`;
      const oxfordUrl = `https://www.oxfordlearnersdictionaries.com/definition/english/${encodedWord}`;
      
      cambridgeFormulas.push([`=HYPERLINK("${cambridgeUrl}", "Cambridge")`]);
      oxfordFormulas.push([`=HYPERLINK("${oxfordUrl}", "Oxford")`]);

      cambridgeColors.push([checkLinkValidity(cambridgeUrl) ? null : BROKEN_LINK_COLOR]);
      oxfordColors.push([checkLinkValidity(oxfordUrl) ? null : BROKEN_LINK_COLOR]);
    } else {
      cambridgeFormulas.push(['']);
      oxfordFormulas.push(['']);
      cambridgeColors.push([null]);
      oxfordColors.push([null]);
    }
  });

  // **MODIFIED**: Hyperlink columns are now P (16) and Q (17).
  sheet.getRange(2, 16, cambridgeFormulas.length, 1).setFormulas(cambridgeFormulas);
  sheet.getRange(2, 17, oxfordFormulas.length, 1).setFormulas(oxfordFormulas);
  sheet.getRange(2, 16, cambridgeColors.length, 1).setBackgrounds(cambridgeColors);
  sheet.getRange(2, 17, oxfordColors.length, 1).setBackgrounds(oxfordColors);

  SpreadsheetApp.getUi().alert('Hyperlinks have been populated and verified for all words.');
}

function showReviewDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ReviewDialog.html')
      .setTitle('Vocabulary Review Session')
      .setWidth(500)
      .setHeight(600);
  SpreadsheetApp.getUi().showModelessDialog(html, html.getTitle());
}

function openScramblePuzzle(word, hint) {
  const template = HtmlService.createTemplateFromFile('ScramblePuzzle');
  template.word = word;
  template.hint = hint;
  const html = template.evaluate()
      .setWidth(400)
      .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Unscramble the Word');
}

function openQuizDialog() {
  const html = HtmlService.createHtmlOutputFromFile('QuizDialog.html')
      .setWidth(500)
      .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Multiple-Choice Quiz');
}


/**
 * **NEW**: This function is called by a temporary trigger to process a word in the background.
 * This allows the UI to remain responsive.
 */
function processWordInBackground() {
  // First, find and delete the trigger that called this function to ensure it only runs once.
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    if (trigger.getHandlerFunction() === 'processWordInBackground') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`inside if for "${trigger}"`)
      // break; // Exit after finding and deleting the first one
    }
  }

  // Now, find the word that was just added (the last row) and process it.
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const word = sheet.getRange(lastRow, 1).getValue();

  if (word) {
    // Call the original processing function
    processNewWord(word, lastRow);
  }
}

/**
 * Adds a new function to the global scope for the review panel to call.
 */
function addNewWord(word) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  
  // Check for duplicates first
  const wordListRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const wordValues = wordListRange.getValues();
  for (let i = 0; i < wordValues.length; i++) {
    if (wordValues[i][0].toString().trim().toLowerCase() === word.toLowerCase()) {
      return `The word "${word}" already exists in the list.`;
    }
  }
  
  // Add the new word to the next empty row
  sheet.getRange(lastRow + 1, 1).setValue(word);
  // processNewWord(word, lastRow + 1);

  ScriptApp.newTrigger('processWordInBackground')
      .timeBased()
      .after(1000) // 1000 milliseconds = 1 second
      .create();

  return `Added "${word}" to your vocabulary list.`;
}


/**
 * Shows the search dialog window.
 */
function showSearchDialog() {
  const html = HtmlService.createHtmlOutputFromFile('SearchDialog.html')
      .setTitle('Search Vocabulary')
      .setWidth(500)
      .setHeight(600);
  SpreadsheetApp.getUi().showModelessDialog(html, html.getTitle());
}

/**
 * Searches for words by prefix (case-insensitive).
 * Called by the client-side JavaScript in SearchDialog.html.
 * @param {string} prefix The search term.
 * @returns {Array<Object>} A list of matching words with their row numbers.
 */
function searchWords(prefix) {
  if (!prefix || prefix.trim().length === 0) {
    return [];
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const wordListRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const wordValues = wordListRange.getValues();
  const lowerCasePrefix = prefix.trim().toLowerCase();
  
  const matches = [];
  wordValues.forEach((row, index) => {
    const word = row[0].toString().trim();
    if (word && word.toLowerCase().startsWith(lowerCasePrefix)) {
      matches.push({
        word: word,
        row: index + 2 // Convert 0-based index to 1-based row number
      });
    }
  });
  return matches.slice(0, 15); // Return a max of 15 matches
}

/**
 * Gets all details for a single word by its row number.
 * Called by the client-side JavaScript in SearchDialog.html.
 * @param {number} rowNumber The row number of the word.
 * @returns {Object} An object containing all details for the word.
 */
function getWordDetails(rowNumber) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error(`Sheet with name "${SHEET_NAME}" not found.`);
    }

    // **MODIFIED**: The range now covers all 24 columns.
    const dataRange = sheet.getRange(rowNumber, 1, 1, 24);
    const rowData = dataRange.getValues()[0];
    // **MODIFIED**: Hyperlinks are now in Columns P and Q.
    const linkFormulas = sheet.getRange(rowNumber, 16, 1, 2).getFormulas()[0];

    const word = rowData[0];
    if (!word) {
      return null;
    }

    const cambridgeFormula = linkFormulas[0];
    const oxfordFormula = linkFormulas[1];
        
    const cambridgeUrl = cambridgeFormula.match(/HYPERLINK\("([^"]+)"/i) ? cambridgeFormula.match(/HYPERLINK\("([^"]+)"/i)[1] : '';
    const oxfordUrl = oxfordFormula.match(/HYPERLINK\("([^"]+)"/i) ? oxfordFormula.match(/HYPERLINK\("([^"]+)"/i)[1] : '';

    // **MODIFIED**: All column indexes updated, and new Persian examples added.
    return {
        word: word,
        row: rowNumber,
        persianTranslations: rowData[1] || '',
        definitions: rowData[2] || '',
        definitionExamples: rowData[3] || '',
        generalExamples: rowData[4] || '',
        generalExamplesPersian: rowData[5] || '',
        partOfSpeech: rowData[6],
        synonyms: rowData[7],
        antonyms: rowData[8],
        notes: rowData[9],
        wordFamily: rowData[10],
        ukPronunciation: rowData[13],
        usPronunciation: rowData[14],
        cambridgeUrl: cambridgeUrl,
        oxfordUrl: oxfordUrl,
        reviewCount: rowData[18] || 0,
        totalReviews: rowData[20] || 0
    };
  } catch (e) {
    Logger.log(`Error in getWordDetails for row ${rowNumber}: ${e.message}`);
    return null;
  }
}

/**
 * Inserts a new, empty column at position F by shifting all subsequent columns to the right.
 * Includes a confirmation dialog because this is a major structural change.
 */
function shiftDataForNewColumn() {
  const ui = SpreadsheetApp.getUi();
  const confirmation = ui.alert(
    'Confirm: Shift Sheet Data',
    'This will insert a new empty column at F and shift all existing data from F to the right. This action cannot be easily undone. Are you sure you want to proceed?',
    ui.ButtonSet.YES_NO
  );

  if (confirmation === ui.Button.YES) {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
      if (!sheet) {
        throw new Error(`Sheet named "${SHEET_NAME}" not found.`);
      }

      // Insert a new column after column E (column number 5).
      // This creates a new, blank column at F.
      sheet.insertColumnAfter(5);
      initializeSheet()
      // Reset column 22 to plain number format
      sheet.getRange(2, 22, sheet.getLastRow()-1).setValue(0);
      sheet.getRange(2, 23, sheet.getLastRow()-1).setValue(0);
      sheet.getRange(2, 24, sheet.getLastRow()-1).setValue(0);
      formatSheet_onlyColor()
      SpreadsheetApp.flush(); // Ensures changes are applied immediately.
      ui.alert( 'A new column has been created at F, and all data has been shifted.');

    } catch (e) {
      Logger.log(e);
      ui.alert('An error occurred: ' + e.message);
    }
  } else {
    ui.alert('Operation cancelled.');
  }
}


// --- BATCH UPDATER FUNCTIONS ---

/**
 * Shows the batch updater sidebar UI.
 */
function showBatchUpdaterSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('BatchUpdaterSidebar.html')
      .setTitle('Batch Word Updater');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Gets the current progress (last processed row) from script properties.
 * Called by the sidebar when it loads.
 */
function getBatchProcessStatus() {
  const properties = PropertiesService.getScriptProperties();
  const lastProcessed = parseInt(properties.getProperty(SCRIPT_PROPERTY_KEY) || '1'); // Default to 1 (header row)
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const totalRows = sheet.getLastRow() - 1;
  
  return {
    lastProcessed: lastProcessed - 1,
    totalRows: totalRows,
    isComplete: (lastProcessed - 1) >= totalRows
  };
}

/**
 * Resets the progress by deleting the script property.
 */
function resetBatchProcess() {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty(SCRIPT_PROPERTY_KEY);
  return getBatchProcessStatus();
}

/**
 * **CORRECTED**: Processes the next batch of words, correctly handling the
 * multi-part-of-speech AI response and preserving user-set data.
 */
function processNextBatch() {
  const properties = PropertiesService.getScriptProperties();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  
  const lastProcessedRow = parseInt(properties.getProperty(SCRIPT_PROPERTY_KEY) || '1');
  const startRow = lastProcessedRow + 1;
  const lastRowInSheet = sheet.getLastRow();

  if (startRow > lastRowInSheet) {
    return getBatchProcessStatus(); // Already done
  }

  const numToProcess = Math.min(BATCH_SIZE, lastRowInSheet - startRow + 1);
  // Read the full width of the sheet to preserve all data
  const sourceRange = sheet.getRange(startRow, 1, numToProcess, 24);
  const sourceData = sourceRange.getValues();
  
  // Prepare all API requests simultaneously
  const requests = sourceData.map(row => prepareGeminiRequest(row[0]));
  const responses = UrlFetchApp.fetchAll(requests);
  const dataToWrite = [];

  // Loop through each API response for each word in the batch
  responses.forEach((response, index) => {
    let newRowData = sourceData[index]; // Start with the original row data to preserve columns
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      try {
        const geminiResponseArray = parseGeminiResponse(response.getContentText());
        
        // Initialize strings to build the new content for each cell
        let allPartsOfSpeech = '', allPersianTranslations = '', allDefinitions = '', allDefinitionExamples = '';
        let allGeneralExamplesEN = '', allGeneralExamplesPER = '', allSynonyms = '', allAntonyms = '';
        let allNotes = '', allWordFamily = '', allUkPronunciation = '', allUsPronunciation = '';

        // Loop through each part of speech object in the array returned by Gemini
        geminiResponseArray.forEach((posData, posIndex) => {
            if (posIndex > 0) {
                const separator = '\n\n';
                allPartsOfSpeech += ', ';
                allPersianTranslations += separator;
                allDefinitions += separator;
                allDefinitionExamples += separator;
                allGeneralExamplesEN += separator;
                allGeneralExamplesPER += separator;
                allSynonyms += separator;
                allAntonyms += separator;
                allNotes += separator;
                allWordFamily += separator;
            }
            const posHeader = `--- ${posData.partOfSpeech.toUpperCase()} ---`;
            allPartsOfSpeech += posData.partOfSpeech || '—';
            allDefinitions += `${posHeader}\n` + (posData.meanings || []).map(m => `• ${m.definition}`).join('\n');
            allPersianTranslations += `${posHeader}\n` + (posData.meanings || []).map(m => `• ${m.persianTranslation}`).join('\n');
            allDefinitionExamples += `${posHeader}\n` + (posData.meanings || []).map(m => `• ${m.example}`).join('\n');
            allGeneralExamplesEN += `${posHeader}\n` + (posData.generalExamples || []).map(item => `• ${item.example}`).join('\n');
            allGeneralExamplesPER += `${posHeader}\n` + (posData.generalExamples || []).map(item => `• ${item.translation}`).join('\n');
            allSynonyms += `${posHeader}\n` + (posData.synonyms || []).map(s => `• ${s}`).join('\n');
            allAntonyms += `${posHeader}\n` + (posData.antonyms || []).map(a => `• ${a}`).join('\n');
            allNotes += `${posHeader}\n` + (posData.notes || []).map(n => `• ${n}`).join('\n');
            allWordFamily += `${posHeader}\n` + (posData.wordFamily || []).map(item => `• ${item.word} (${item.POS})`).join('\n');
            allUkPronunciation += `${posData.partOfSpeech}:${posData.ukPronunciation || '—'}\n`;
            allUsPronunciation += `${posData.partOfSpeech}:${posData.usPronunciation || '—'}\n`;
        });

        // --- Map the new, formatted strings back into the row data array ---
        // Indexes are 0-based, so Column B is 1, C is 2, etc.
        newRowData[1] = allPersianTranslations.trim();
        newRowData[2] = allDefinitions.trim();
        newRowData[3] = allDefinitionExamples.trim();
        newRowData[4] = allGeneralExamplesEN.trim();
        newRowData[5] = allGeneralExamplesPER.trim();
        newRowData[6] = allPartsOfSpeech.trim();
        newRowData[7] = allSynonyms.trim();
        newRowData[8] = allAntonyms.trim();
        newRowData[9] = allNotes.trim();
        newRowData[10] = allWordFamily.trim();
        // Column L (index 11 - Created At) is preserved from the original data
        newRowData[12] = new Date(); // Column M (index 12) - Modified At is updated
        newRowData[13] = allUkPronunciation.trim(); // Column N
        newRowData[14] = allUsPronunciation.trim(); // Column O
        // Columns P & Q (15 & 16) for links are preserved
        // Columns R, S, T, U (17-20) for review stats are preserved
        // Columns V, W, X (21-23) for metadata are preserved

      } catch(e) {
        Logger.log(`Error parsing response for word "${sourceData[index][0]}": ${e.message}`);
      }
    } else {
      Logger.log(`API error for word "${sourceData[index][0]}": HTTP ${responseCode}`);
    }
    dataToWrite.push(newRowData);
  });

  // Write the entire updated batch back to the sheet
  sourceRange.setValues(dataToWrite);
  
  // Update and save the progress
  const newLastProcessedRow = startRow + numToProcess - 1;
  properties.setProperty(SCRIPT_PROPERTY_KEY, newLastProcessedRow.toString());

  return getBatchProcessStatus();
}




/**
 * Adds new options to the custom menu.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Review Vocabulary')
      .addItem('Start Review Session', 'showReviewDialog')
      // .addItem('Start Multiple-Choice Quiz', 'openQuizDialog')
      .addItem('Search Words', 'showSearchDialog') // <-- ADD THIS LINE
      .addSeparator()
      .addItem('Format Sheet', 'formatSheet')
      .addItem('Remove Duplicates', 'removeDuplicateWords')
      .addSeparator()
      .addItem('Initialize Sheet', 'initializeSheet')
      .addItem('Initialize Quiz Counts', 'initializeQuizCounts')
      .addItem('Initialize Total Reviews', 'initializeTotalReviews')
      .addItem('Populate & Verify Hyperlinks', 'populateHyperlinks')
      .addItem('Schedule Initial Reviews', 'scheduleInitialReviews')
      .addSeparator()
      .addItem('SHIFT DATA for New Column', 'shiftDataForNewColumn') // <-- ADD THIS LINE
      .addSeparator()
      .addItem('Batch Update Words', 'showBatchUpdaterSidebar')
      .addToUi();
}
