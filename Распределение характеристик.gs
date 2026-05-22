// ЛОГИКА РАСПРЕДЕЛЕНИЯ ОЧКОВ
function mainProcess(sheet) {
  const cond1 = sheet.getRange("AG26").getValue();
  const cond2 = sheet.getRange("AG48").getValue();
  const warningCell = sheet.getRange("AG5");

  if (cond1 === 0 && cond2 === 0) {
    warningCell.clearContent();
    const blocks = [{range: "AH9:AH23", source: "AI9:AI23"}, {range: "AH35:AH46", source: "AI35:AI46"}];
    
    blocks.forEach(b => {
      const rDest = sheet.getRange(b.range);
      const rSrc = sheet.getRange(b.source);
      const valsSrc = rSrc.getValues();
      const result = rDest.getValues().map((row, i) => [(Number(row) || 0) + (Number(valsSrc[i]) || 0)]);
      rDest.setValues(result);
      rSrc.clearContent();
    });
  } else {
    warningCell.setValue("Распределите все очки прежде чем продолжить");
    SpreadsheetApp.flush();
    Utilities.sleep(3000);
    warningCell.clearContent();
  }
}
