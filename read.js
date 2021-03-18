const workbook = XLSX.readFile("data.js");
let worksheets = {};
for (const sheetName of workbook.SheetNames) {
  // Some helper functions in XLSX.utils generate different views of the sheets:
  //     XLSX.utils.sheet_to_csv generates CSV
  //     XLSX.utils.sheet_to_txt generates UTF16 Formatted Text
  //     XLSX.utils.sheet_to_html generates HTML
  //     XLSX.utils.sheet_to_json generates an array of objects
  //     XLSX.utils.sheet_to_formulae generates a list of formulae
  worksheets[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
}

// Show the data as JSON
console.log("json:\n", JSON.stringify(worksheets.Sheet1), "\n\n");
console.log(
  "xml:\n",
  jsontoxml(
    {
      worksheets: JSON.parse(JSON.stringify(Object.values(worksheets))).map(
        (worksheet) =>
          worksheet.map((data) => {
            for (property in data) {
              const newPropertyName = property.replace(/\s/g, "");
              if (property !== newPropertyName) {
                Object.defineProperty(
                  data,
                  newPropertyName,
                  Object.getOwnPropertyDescriptor(data, property)
                );
                delete data[property];
              }
            }
            return data;
          })
      ),
    },
    {}
  ),
  "\n\n"
);

// Modify the XLSX
worksheets.Sheet1.push();
