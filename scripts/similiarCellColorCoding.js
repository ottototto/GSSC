function colorCodeSelectedBasedOnSimilarity() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();
  const threshold = 0.5; // Adjust the threshold as needed
  const colorMap = new Map(); // Map to store text to color mapping
  const colors = ['#FFCCCC', '#CCFFCC', '#CCCCFF', '#FFFF99', '#FFCC99', '#99CCFF', '#CC99FF']; // Color palette, change as you wish.
  let colorIndex = 0;

  let flatValues = [];
  values.forEach((row, rowIndex) => {
    row.forEach((cellValue, colIndex) => {
      if (cellValue.trim() !== '') { // Ignore empty cells
        flatValues.push({
          value: cellValue,
          row: rowIndex + range.getRow(), // Adjust for actual position in sheet
          col: colIndex + range.getColumn(),
        });
      }
    });
  });

  function findOrAssignColor(text) {
    for (let [key, value] of colorMap) {
      if (calculateSimilarity(key, text) >= threshold) {
        return value; // Return existing color if similar enough
      }
    }
    const color = colors[colorIndex++ % colors.length];
    colorMap.set(text, color);
    return color;
  }

  flatValues.forEach(cell => {
    const color = findOrAssignColor(cell.value);
    sheet.getRange(cell.row, cell.col).setBackground(color);
  });
}

function calculateSimilarity(text1, text2) {
  const words1 = text1.toLowerCase().split(/\s+/);
  const words2 = text2.toLowerCase().split(/\s+/);
  const setWords1 = new Set(words1);
  const setWords2 = new Set(words2);
  const commonWords = [...setWords1].filter(word => setWords2.has(word));
  const similarity = 2 * commonWords.length / (words1.length + words2.length);
  return similarity;
}