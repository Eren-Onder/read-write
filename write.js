// Create a new XLSX file
const newBook = XLSX.utils.book_new();
const newSheet = XLSX.utils.json_to_sheet(worksheets.Sheet1);
XLSX.utils.book_append_sheet(newBook, newSheet, "Sheet1");
XLSX.writeFile(newBook, "eren.xlsx");
