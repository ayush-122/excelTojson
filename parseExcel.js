const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

// Load the Excel file
const filePath = path.join("C:", "Users", "AyushSingh", "Downloads", "MilkyWays_Combo_DRAFT.xlsx"); // Replace with your file path
const workbook = xlsx.readFile(filePath);

// Filter sheets whose names start with "export_excel"
const exportSheets = workbook.SheetNames.filter((name) =>
  name.startsWith("export")
);

// Define parsing rules for each tag
const TAG_PARSING_RULES = {
  default: "default", // Fallback parsing rule
  paylines: "rowByRow", // Custom parsing logic for #paylines
  settings: "keyValue", // Example for key-value parsing
  paytable: "keyValue",
  metadata: "customFunction", // Custom processing via a function
  weight: "weightRowByRow" // Custom parsing logic for tags starting with WEIGHT
};

// Helper function to filter valid data from rows
const filterValidValues = (row) =>
  row.filter((value) => value !== null && value !== "" && value !== undefined);

// Parsing logic for each type of tag
const parseDataByTag = (tagName, headers, dataRows) => {
  // Determine the parsing rule dynamically based on the tag name
  const parsingRule = tagName.toLowerCase().startsWith("weight")
    ? TAG_PARSING_RULES.weight
    : TAG_PARSING_RULES[tagName.toLowerCase()] || TAG_PARSING_RULES.default;

  // Filter out completely empty rows
  const validDataRows = dataRows.filter((row) =>
    row.some((value) => value !== null && value !== "" && value !== undefined)
  );

  switch (parsingRule) {
    case "weightRowByRow":
      // Special parsing logic for WEIGHT tags
      return parseWeightRows(dataRows, headers);

    case "rowByRow":
      // Parse each row as an array of valid values
      return validDataRows.map((row) => {
        if (row.every((cell) => cell === null || cell === "" || cell === undefined)) {
          return null; // End the section if the row is empty
        }
        return filterValidValues(row);
      }).filter(item => item !== null); // Remove null items (empty rows)

    case "keyValue":
      // Parse rows as key-value pairs (first column is the key)
      return validDataRows.reduce((obj, row) => {
        if (row.every((cell) => cell === null || cell === "" || cell === undefined)) {
          return obj; // No more data to process
        }

        const filteredRow = filterValidValues(row);
        const key = filteredRow[0]; // First column is the key
        const value = filteredRow.slice(1); // Remaining columns are the value
        if (key) obj[key] = value;
        return obj;
      }, {});

    case "customFunction":
      // Custom parsing logic
      return customParsingFunction(headers, validDataRows);

    case "default":
    default:
      // Default parsing: Map headers to data columns
      return headers.reduce((obj, header, idx) => {
        obj[header] = validDataRows
          .map((row) => {
            if (row.every((cell) => cell === null || cell === "" || cell === undefined)) {
              return null;
            }
            return row[idx];
          })
          .filter((value) => value !== null && value !== "" && value !== undefined); // Remove invalid values
        return obj;
      }, {});
  }
};

// Function to parse weight rows
const parseWeightRows = (rows, headers) => {
  //console.log("inside parseWeightRows", rows, headers);

  const result = []; // Array to hold the final parsed objects

  // Iterate over rows
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex];
    if (row.length === 0 || row[0] === undefined) break;

    const obj = {};

    // Iterate over headers
    for (let idx = 0; idx < headers.length; idx++) {
      const header = headers[idx];

      // Check if the header is empty or invalid
      if (!header || header.length === 0 || row[idx] === undefined) {
        break; // Skip processing for this header
      }

      obj[header] = row[idx] !== undefined ? row[idx] : "";
    }

    // Add the constructed object to the result array
    result.push(obj);
  }

  return result; // Return the final parsed array
};

// Example custom function (for tags like #metadata)
const customParsingFunction = (headers, dataRows) => {
  return dataRows.map((row) =>
    headers.reduce((obj, header, idx) => {
      obj[header] = row[idx];
      return obj;
    }, {})
  );
};

// Initialize an object to store data from filtered sheets
let data = {}; // This should be declared outside the loop

// Process each sheet
workbook.SheetNames.forEach((sheetName) => {
  const sheet = workbook.Sheets[sheetName];

  // Convert sheet to JSON (2D array format)
  const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });

  rows.forEach((row, rowIndex) => {
    // Step 1: Detect tags in the row
    const tags = row
      .map((cell, colIndex) =>
        typeof cell === "string" && cell.startsWith("#")
          ? { tag: cell.replace("#", ""), colIndex }
          : null
      )
      .filter((item) => item !== null);

    if (tags.length > 0) {
      // Step 2: Process each tag and its corresponding columns
      tags.forEach((tag, tagIndex) => {
        const { tag: tagName, colIndex } = tag;

        // Skip processing for the "end" tag
        if (tagName.toLowerCase() === "end") {
          return;
        }

        if (!data[tagName]) {
          data[tagName] = []; // Initialize tag array
        }

        // Step 3: Determine column range for the tag
        let nextTagColIndex;
        if (tagName.toLowerCase().startsWith("weight")) {
          nextTagColIndex = colIndex+1; // For WEIGHT tags, only process the next column
        } else {
          nextTagColIndex =
            tags[tagIndex + 1]?.colIndex ||
            tags.find((t) => t.tag === "end")?.colIndex ||
            row.length + 1; // Default behavior
        }

        const headersRow = rows[rowIndex + 1] || []; // Header row is the next row or empty if missing
        const dataRows = rows.slice(rowIndex + 2); // Data rows start after headers

        const headers = headersRow
          .slice(colIndex, nextTagColIndex + 1)
          .map((header, idx) =>
            typeof header === "string"
              ? header.toLowerCase()
              : `column_${colIndex + idx}` // Fallback to generic column name
          );

        // Use parsing logic based on the tag
        data[tagName] = parseDataByTag(
          tagName,
          headers,
          dataRows.map((row) => row.slice(colIndex, nextTagColIndex + 1))
        );
      });
    }
  });
});

// Save JSON output to a file
const jsonOutput = JSON.stringify(data, null, 4);
fs.writeFileSync("output.json", jsonOutput);

console.log("JSON saved to output.json");
