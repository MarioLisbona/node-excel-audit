// Helper function to convert a column index to a letter (e.g., 0 -> A, 4 -> E, 26 -> AA)
export function getColumnLetter(colIndex) {
  let letter = "";
  while (colIndex >= 0) {
    letter = String.fromCharCode((colIndex % 26) + 65) + letter;
    colIndex = Math.floor(colIndex / 26) - 1;
  }
  return letter;
}
