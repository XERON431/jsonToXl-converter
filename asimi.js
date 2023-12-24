const excel = require('excel4node');
const fs = require('fs');

// Read data from input.json
const inputData = JSON.parse(fs.readFileSync('input.json'));

// Create a new instance of a Workbook class
const wb = new excel.Workbook();

// Object to store references to sheets
const sheetReferences = {};

// Object to store created worksheets for easy reference
const createdWorksheets = {};

function createWorksheets(obj, sheetName, parentSheet) {
  // Create a worksheet
  const ws = wb.addWorksheet(sheetName);
  createdWorksheets[sheetName] = ws;

  let colNum = 1;
  for (const key in obj) {
    if (Array.isArray(obj[key]) && key !== 'sections') {
      // Handling arrays
      ws.cell(1, colNum).string(key);
      ws.cell(2, colNum).string(obj[key].join(', '));
      colNum++;
    } else if (typeof obj[key] === 'object' && obj[key] !== null) {
      if (key === 'sections') {
        // Handling 'sections' array of objects
        obj[key].forEach((section, index) => {
          const sectionName = section.sectionName;
          const sectionSheetName = `${sheetName}_${sectionName}`;
          const sectionWs = wb.addWorksheet(sectionSheetName);
          createWorksheets(section, sectionSheetName, sheetName);

          // Handling 'books' array within 'sections'
          let sectionColNum = 1;
          const books = section.books;
          if (Array.isArray(books) && books.length > 0) {
            const bookKeys = Object.keys(books[0]);
            sectionWs.cell(1, 1).string('Book');
            bookKeys.forEach((bookKey, i) => {
              sectionWs.cell(1, i + 2).string(bookKey);
              books.forEach((book, j) => {
                sectionWs.cell(j + 2, 1).string(`Book ${j + 1}`);
                sectionWs.cell(j + 2, i + 2).string(String(book[bookKey]));
              });
            });
          }
        });
      } else {
        // Handling nested objects
        const nestedSheetName = `${sheetName}_${key}`;
        ws.cell(1, colNum).string(key);
        ws.cell(2, colNum).string(nestedSheetName);
        createWorksheets(obj[key], nestedSheetName, sheetName);

        // Storing references of nested sheets
        if (!sheetReferences[parentSheet]) {
          sheetReferences[parentSheet] = [];
        }
        sheetReferences[parentSheet].push(nestedSheetName);

        colNum++;
      }
    } else {
      // Handling simple values
      ws.cell(1, colNum).string(key);
      ws.cell(2, colNum).string(String(obj[key]));
      colNum++;
    }
  }
}

// Start creating worksheets from the main input data
createWorksheets(inputData, 'Main');

// Writing references for parent sheets
Object.keys(sheetReferences).forEach(parentSheet => {
  const col = 10; // Column number to place references
  const parentSheetReferences = sheetReferences[parentSheet];
  if (parentSheetReferences) {
    const parentWs = createdWorksheets[parentSheet];
    if (parentWs) {
      const lastRow = parentWs.rowCount + 1;
      parentSheetReferences.forEach((sheet, index) => {
        parentWs.cell(lastRow + index, col).string(sheet);
      });
    }
  }
});

// Writing the Excel file
wb.write('Excel.xlsx', (err, stats) => {
  if (err) {
    console.error(err);
    return;
  }
  console.log('Excel file created successfully!');
});
