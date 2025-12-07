/**
 * Building Structure Generator - Graph Paper Mode (V6: Floor Numbers)
 * - Added Floor Numbers (FL 1, FL 2...) on the right side of Side View
 */

const CONFIG = {
  sheetInput: "Input",
  sheetPlan: "Plan",
  cellSizePx: 12,       // ขนาดเซลล์
  resolution: 0.5,      // 1 ช่อง = 0.5 เมตร
  startRow: 6,          // เริ่มวาดห่างจากขอบบน
  startCol: 6,          // เริ่มวาดห่างจากขอบซ้าย
  stumpHeight: 2,       // ความสูงเสาตอม่อ
  colors: {
    beam: "#37474f",      // สีโครงสร้าง
    fillSide: "#f3e5f5",  // สีพื้น Side View
    fillTop: "#e1f5fe",   // สีพื้น Top View
    gridLabel: "#b71c1c", // สีตัวหนังสือ Grid
    dimText: "#0d47a1",   // สีตัวเลขบอกระยะ
    graphLine: "#eceff1", // สีเส้นตารางกราฟ
    labelBg: "#ffffff",   // สีพื้นป้าย Label
    support: "#424242"    // สีฐานราก
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏗️ Building Tools')
    .addItem('Create/Reset Input Sheet', 'setupInputSheet')
    .addItem('Generate Blueprint', 'generateBlueprint')
    .addToUi();
}

function setupInputSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.sheetInput);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.sheetInput);
  } else {
    sheet.clear();
  }

  const headers = [
    ["PARAMETER", "VALUE (CSV)", "DESCRIPTION"],
    ["Span X (m)", "4,4,4", "ระยะห่างเสาแนวนอน (คั่นด้วยจุลภาค)"],
    ["Height (m)", "3.5,3.5,3.5", "ความสูงแต่ละชั้น (จากพื้นขึ้นบน)"],
    ["Span Y (m)", "4,3,4", "ระยะห่างเสาแนวลึก (Top View)"]
  ];

  sheet.getRange("A1:C4").setValues(headers);
  sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#eeeeee");
  sheet.getRange("A1:C4").setBorder(true, true, true, true, true, true);
  sheet.autoResizeColumns(1, 3);
}

function createLabelBox(sheet, row, col, text, color, align = "center", isBold = true) {
  if (row < 1 || col < 1) return; 
  const range = sheet.getRange(row, col, 2, 2); 
  range.merge();
  range.setValue(text);
  range.setFontColor(color);
  range.setBackground(CONFIG.colors.labelBg);
  range.setBorder(true, true, true, true, null, null, "#dddddd", SpreadsheetApp.BorderStyle.SOLID);
  if (isBold) range.setFontWeight("bold");
  range.setHorizontalAlignment(align).setVerticalAlignment("middle");
  range.setFontSize(8);
}

function drawFixedSupport(sheet, row, centerX) {
  const width = 4; 
  const height = 3; 
  const startX = centerX - Math.floor(width / 2);
  if (startX < 1) return;

  const range = sheet.getRange(row, startX, height, width);
  range.merge();
  range.setBackground(CONFIG.colors.support);
  range.setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function drawColumnStump(sheet, row, x, height) {
  sheet.getRange(row, x, height, 1)
       .setBorder(null, true, null, null, null, null, CONFIG.colors.beam, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function generateBlueprint() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(CONFIG.sheetInput);
  if (!inputSheet) { SpreadsheetApp.getUi().alert("Please setup input first."); return; }

  const rawSpanX = inputSheet.getRange("B2").getValue();
  const rawHeight = inputSheet.getRange("B3").getValue();
  const rawSpanY = inputSheet.getRange("B4").getValue();

  if (!rawSpanX || !rawHeight || !rawSpanY) return;

  const parseToCells = (str) => str.toString().split(',').map(n => Math.round(parseFloat(n) / CONFIG.resolution));
  const parseToMeters = (str) => str.toString().split(',').map(Number);
  
  const spansX_cells = parseToCells(rawSpanX);
  const spansX_meters = parseToMeters(rawSpanX);
  const heights_cells = parseToCells(rawHeight).reverse(); 
  const heights_meters = parseToMeters(rawHeight).reverse();
  const spansY_cells = parseToCells(rawSpanY);
  const spansY_meters = parseToMeters(rawSpanY);

  const totalWidthCells = spansX_cells.reduce((a, b) => a + b, 0) + CONFIG.startCol + 20; 
  const totalHeightCells_Side = heights_cells.reduce((a, b) => a + b, 0);
  const totalHeightCells_Top = spansY_cells.reduce((a, b) => a + b, 0);
  const totalRowsNeeded = totalHeightCells_Side + totalHeightCells_Top + CONFIG.startRow + 50; 

  let planSheet = ss.getSheetByName(CONFIG.sheetPlan);
  if (planSheet) ss.deleteSheet(planSheet);
  planSheet = ss.insertSheet(CONFIG.sheetPlan);

  if (totalWidthCells > planSheet.getMaxColumns()) planSheet.insertColumnsAfter(planSheet.getMaxColumns(), totalWidthCells - planSheet.getMaxColumns());
  if (totalRowsNeeded > planSheet.getMaxRows()) planSheet.insertRowsAfter(planSheet.getMaxRows(), totalRowsNeeded - planSheet.getMaxRows());

  planSheet.setColumnWidths(1, totalWidthCells, CONFIG.cellSizePx);
  planSheet.setRowHeights(1, totalRowsNeeded, CONFIG.cellSizePx);
  
  planSheet.getRange(1, 1, totalRowsNeeded, totalWidthCells)
    .setBorder(true, true, true, true, true, true, CONFIG.colors.graphLine, SpreadsheetApp.BorderStyle.DOTTED);

  // ==========================================
  // DRAW SIDE VIEW
  // ==========================================
  let currentRow = CONFIG.startRow;
  let startCol = CONFIG.startCol;

  planSheet.getRange(currentRow - 4, startCol).setValue("SIDE VIEW (Elevation)").setFontSize(12).setFontWeight("bold");

  heights_cells.forEach((hCells, index) => {
    let currentX = startCol;
    const hMeters = heights_meters[index];
    
    // Label ความสูง (ซ้าย)
    createLabelBox(planSheet, currentRow + Math.floor(hCells/2) - 1, startCol - 3, `${hMeters}m`, CONFIG.colors.dimText);

    // วาดห้อง
    for (let i = 0; i < spansX_cells.length; i++) {
      const wCells = spansX_cells[i];
      const room = planSheet.getRange(currentRow, currentX, hCells, wCells);
      room.setBackground(CONFIG.colors.fillSide);
      room.setBorder(true, true, true, true, null, null, CONFIG.colors.beam, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      currentX += wCells;
    }
    
    // === [เพิ่มใหม่] Label เลขชั้น (ขวา) ===
    // currentX ตอนนี้อยู่ที่ขอบขวาสุดของอาคารแล้ว
    const floorNum = heights_cells.length - index; // คำนวณเลขชั้น (เช่น 3, 2, 1)
    createLabelBox(planSheet, currentRow + Math.floor(hCells/2) - 1, currentX + 1, `FL ${floorNum}`, CONFIG.colors.gridLabel);

    currentRow += hCells;
  });

  // DRAW STUMP & SUPPORT
  let columnX = startCol;
  const colPositions = [startCol];
  spansX_cells.forEach(w => {
    columnX += w;
    colPositions.push(columnX);
  });

  colPositions.forEach(x => {
    drawColumnStump(planSheet, currentRow, x, CONFIG.stumpHeight);
    drawFixedSupport(planSheet, currentRow + CONFIG.stumpHeight, x);
  });

  // GRID LABEL & DIM (Side View)
  let gridX = startCol;
  let gridNum = 1;
  const labelRow = currentRow + CONFIG.stumpHeight + 5; 

  createLabelBox(planSheet, labelRow, gridX - 1, gridNum++, CONFIG.colors.gridLabel);

  for (let i = 0; i < spansX_cells.length; i++) {
    const wCells = spansX_cells[i];
    const wMeters = spansX_meters[i];
    createLabelBox(planSheet, labelRow, gridX + Math.floor(wCells/2) - 1, `${wMeters}m`, CONFIG.colors.dimText, "center", false);
    gridX += wCells;
    createLabelBox(planSheet, labelRow, gridX - 1, gridNum++, CONFIG.colors.gridLabel);
  }

  // ==========================================
  // DRAW TOP VIEW
  // ==========================================
  currentRow += 15;
  planSheet.getRange(currentRow - 4, startCol).setValue("TOP VIEW (Plan)").setFontSize(12).setFontWeight("bold");

  gridX = startCol;
  gridNum = 1;
  createLabelBox(planSheet, currentRow - 3, gridX - 1, gridNum++, CONFIG.colors.gridLabel);

  for (let i = 0; i < spansX_cells.length; i++) {
    const wCells = spansX_cells[i];
    const wMeters = spansX_meters[i];
    createLabelBox(planSheet, currentRow - 3, gridX + Math.floor(wCells/2) - 1, `${wMeters}m`, CONFIG.colors.dimText, "center", false);
    gridX += wCells;
    createLabelBox(planSheet, currentRow - 3, gridX - 1, gridNum++, CONFIG.colors.gridLabel);
  }

  spansY_cells.forEach((hCells, index) => {
    let currentX = startCol;
    const hMeters = spansY_meters[index];
    const charCode = 65 + index;

    createLabelBox(planSheet, currentRow - 1, startCol - 3, String.fromCharCode(charCode), CONFIG.colors.gridLabel);
    createLabelBox(planSheet, currentRow + Math.floor(hCells/2) - 1, startCol - 3, `${hMeters}m`, CONFIG.colors.dimText, "center", false);

    for (let i = 0; i < spansX_cells.length; i++) {
      const wCells = spansX_cells[i];
      const room = planSheet.getRange(currentRow, currentX, hCells, wCells);
      room.setBackground(CONFIG.colors.fillTop);
      room.setBorder(true, true, true, true, null, null, "#90a4ae", SpreadsheetApp.BorderStyle.SOLID);
      currentX += wCells;
    }
    
    if (index === spansY_cells.length - 1) {
       createLabelBox(planSheet, currentRow + hCells - 1, startCol - 3, String.fromCharCode(charCode + 1), CONFIG.colors.gridLabel);
    }

    currentRow += hCells;
  });
  
  planSheet.setHiddenGridlines(true);
}