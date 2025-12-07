/**
 * Building Structure Generator - Sidebar Edition (V8: Centered & Cleanup)
 * - Auto-delete Input sheet
 * - Auto-center drawing in the middle of the sheet
 */

const CONFIG = {
  sheetPlan: "Plan",
  cellSizePx: 12,       // 1 Cell Pixel Size
  resolution: 0.5,      // 1 Cell = 0.5m
  minPadding: 10,       // ระยะขอบต่ำสุด (จำนวนช่อง)
  stumpHeight: 2,
  colors: {
    beam: "#37474f",
    fillSide: "#f3e5f5",
    fillTop: "#e1f5fe",
    gridLabel: "#b71c1c",
    dimText: "#0d47a1",
    graphLine: "#eceff1",
    labelBg: "#ffffff",
    support: "#424242"
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏗️ CVE-RU ELF')
    .addItem('START ELF Program', 'showSidebar')
    .addToUi();

  // สั่งเปิด Sidebar อัตโนมัติทันที
  showSidebar();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Building Generator')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html); // ใช้ Sidebar เหมือนเดิม
}

function receiveFormInput(spanXStr, heightStr, spanYStr) {
  generateBlueprintFromData(spanXStr, heightStr, spanYStr);
}

// --- CORE LOGIC ---
function generateBlueprintFromData(rawSpanX, rawHeight, rawSpanY) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Cleanup: ลบหน้า Input ทิ้ง (ถ้ามี)
  const inputSheet = ss.getSheetByName("Input");
  if (inputSheet) {
    try {
      ss.deleteSheet(inputSheet);
    } catch (e) {
      // กรณีเหลือ Sheet เดียวจะลบไม่ได้ ก็ปล่อยผ่านไปก่อน
    }
  }

  // 2. Parse Inputs
  const parseToCells = (str) => str.toString().split(',').map(n => Math.round(parseFloat(n) / CONFIG.resolution));
  const parseToMeters = (str) => str.toString().split(',').map(Number);
  
  const spansX_cells = parseToCells(rawSpanX);
  const spansX_meters = parseToMeters(rawSpanX);
  const heights_cells = parseToCells(rawHeight).reverse(); 
  const heights_meters = parseToMeters(rawHeight).reverse();
  const spansY_cells = parseToCells(rawSpanY);
  const spansY_meters = parseToMeters(rawSpanY);

  // 3. คำนวณขนาดและจัดกึ่งกลาง
  const drawingWidth = spansX_cells.reduce((a, b) => a + b, 0);
  const totalHeightCells_Side = heights_cells.reduce((a, b) => a + b, 0);
  const totalHeightCells_Top = spansY_cells.reduce((a, b) => a + b, 0);
  
  // กำหนดความกว้างกระดาษ (Canvas) ให้กว้างกว่ารูป + Padding
  // สมมติหน้าจอมาตรฐานแสดงได้ประมาณ 100-150 ช่อง หรือปรับตามรูป
  const canvasWidth = Math.max(drawingWidth + (CONFIG.minPadding * 2), 80); // อย่างน้อย 80 ช่อง
  const totalRowsNeeded = totalHeightCells_Side + totalHeightCells_Top + 50; 

  // คำนวณจุดเริ่มต้น (startCol) ให้รูปอยู่กลาง Canvas
  let startCol = Math.floor((canvasWidth - drawingWidth) / 2);
  if (startCol < 4) startCol = 4; // กันเหนียวไม่ให้ชิดซ้ายเกินไป (เผื่อ Label)
  
  let startRow = 6; // เริ่มวาดบรรทัดที่ 6

  // 4. Setup Sheet Plan
  let planSheet = ss.getSheetByName(CONFIG.sheetPlan);
  if (planSheet) ss.deleteSheet(planSheet);
  planSheet = ss.insertSheet(CONFIG.sheetPlan);

  // Resize Area
  if (canvasWidth > planSheet.getMaxColumns()) planSheet.insertColumnsAfter(planSheet.getMaxColumns(), canvasWidth - planSheet.getMaxColumns());
  // ถ้า Sheet กว้างเกินไป ให้ลบคอลัมน์ส่วนเกินออก เพื่อให้ Scrollbar อยู่ตรงกลางพอดี
  if (planSheet.getMaxColumns() > canvasWidth) {
     planSheet.deleteColumns(canvasWidth + 1, planSheet.getMaxColumns() - canvasWidth);
  }

  if (totalRowsNeeded > planSheet.getMaxRows()) planSheet.insertRowsAfter(planSheet.getMaxRows(), totalRowsNeeded - planSheet.getMaxRows());

  // Set Grid Size
  planSheet.setColumnWidths(1, canvasWidth, CONFIG.cellSizePx);
  planSheet.setRowHeights(1, totalRowsNeeded, CONFIG.cellSizePx);
  
  // Draw Graph Grid
  planSheet.getRange(1, 1, totalRowsNeeded, canvasWidth)
    .setBorder(true, true, true, true, true, true, CONFIG.colors.graphLine, SpreadsheetApp.BorderStyle.DOTTED);

  // ==========================================
  // DRAW SIDE VIEW (Centered)
  // ==========================================
  let currentRow = startRow;
  
  // Header Side View
  planSheet.getRange(currentRow - 4, startCol).setValue("SIDE VIEW (Elevation)").setFontSize(12).setFontWeight("bold");

  heights_cells.forEach((hCells, index) => {
    let currentX = startCol;
    const hMeters = heights_meters[index];
    
    // Level Label (Left)
    createLabelBox(planSheet, currentRow + Math.floor(hCells/2) - 1, startCol - 3, `${hMeters}m`, CONFIG.colors.dimText);

    // Draw Rooms
    for (let i = 0; i < spansX_cells.length; i++) {
      const wCells = spansX_cells[i];
      const room = planSheet.getRange(currentRow, currentX, hCells, wCells);
      room.setBackground(CONFIG.colors.fillSide);
      room.setBorder(true, true, true, true, null, null, CONFIG.colors.beam, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      currentX += wCells;
    }
    
    // Floor Label (Right)
    const floorNum = heights_cells.length - index;
    createLabelBox(planSheet, currentRow + Math.floor(hCells/2) - 1, currentX + 1, `FL ${floorNum}`, CONFIG.colors.gridLabel);

    currentRow += hCells;
  });

  // STUMP & SUPPORT
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

  // SIDE VIEW LABELS (Grid & Dim)
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
  // DRAW TOP VIEW (Centered)
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

// --- HELPER FUNCTIONS ---
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

function generateBlueprintFromData(rawSpanX, rawHeight, rawSpanY) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // === [แก้ไข] ใส่ค่า Default ป้องกัน Error กรณีค่าว่าง ===
  if (!rawSpanX) rawSpanX = "4,4,4";      // ค่าทดสอบ
  if (!rawHeight) rawHeight = "3.5,3.5";  // ค่าทดสอบ
  if (!rawSpanY) rawSpanY = "4,3";        // ค่าทดสอบ
  // =================================================

  // 1. Parse Inputs
  // (โค้ดเดิม) แก้ให้ปลอดภัยขึ้นด้วย (str || "")
  const parseToCells = (str) => (str || "").toString().split(',').map(n => Math.round(parseFloat(n) / CONFIG.resolution));
  const parseToMeters = (str) => (str || "").toString().split(',').map(Number);
  
  const spansX_cells = parseToCells(rawSpanX);
  const spansX_meters = parseToMeters(rawSpanX);
  const heights_cells = parseToCells(rawHeight).reverse(); 
  const heights_meters = parseToMeters(rawHeight).reverse();
  const spansY_cells = parseToCells(rawSpanY);
  const spansY_meters = parseToMeters(rawSpanY);

  // 2. คำนวณขนาดและจัดกึ่งกลาง
  const drawingWidth = spansX_cells.reduce((a, b) => a + b, 0);
  const totalHeightCells_Side = heights_cells.reduce((a, b) => a + b, 0);
  const totalHeightCells_Top = spansY_cells.reduce((a, b) => a + b, 0);
  
  const canvasWidth = Math.max(drawingWidth + (CONFIG.minPadding * 2), 80); 
  const totalRowsNeeded = totalHeightCells_Side + totalHeightCells_Top + 50; 

  let startCol = Math.floor((canvasWidth - drawingWidth) / 2);
  if (startCol < 4) startCol = 4; 
  let startRow = 6;

  // 3. Setup Sheet Plan (แบบปลอดภัย: ไม่ลบจนกว่าจะมีอันใหม่)
  const oldPlan = ss.getSheetByName(CONFIG.sheetPlan);
  if (oldPlan) {
    // เปลี่ยนชื่ออันเก่าหนีก่อน เพื่อให้ชื่อ "Plan" ว่างสำหรับอันใหม่
    oldPlan.setName("Plan_Old_Deleting");
  }

  // สร้าง Sheet ใหม่ (ตอนนี้จะมี 2 Sheet คือ Plan_Old_Deleting กับ Plan)
  const planSheet = ss.insertSheet(CONFIG.sheetPlan);

  // ตอนนี้ปลอดภัยแล้ว ลบอันเก่าทิ้งได้ (เพราะมี planSheet ใหม่รองรับแล้ว)
  if (oldPlan) {
    ss.deleteSheet(oldPlan);
  }

  // ลบหน้า Input ทิ้ง (ถ้ามี)
  const inputSheet = ss.getSheetByName("Input");
  if (inputSheet) {
    try {
      ss.deleteSheet(inputSheet);
    } catch (e) {
      // ถ้าลบไม่ได้ (กรณีแปลกๆ) ก็ปล่อยไว้ ไม่ซีเรียส
    }
  }

  // 4. Resize Area (ทำกับ Sheet ใหม่)
  if (canvasWidth > planSheet.getMaxColumns()) planSheet.insertColumnsAfter(planSheet.getMaxColumns(), canvasWidth - planSheet.getMaxColumns());
  if (planSheet.getMaxColumns() > canvasWidth) {
     planSheet.deleteColumns(canvasWidth + 1, planSheet.getMaxColumns() - canvasWidth);
  }
  if (totalRowsNeeded > planSheet.getMaxRows()) planSheet.insertRowsAfter(planSheet.getMaxRows(), totalRowsNeeded - planSheet.getMaxRows());

  planSheet.setColumnWidths(1, canvasWidth, CONFIG.cellSizePx);
  planSheet.setRowHeights(1, totalRowsNeeded, CONFIG.cellSizePx);
  
  planSheet.getRange(1, 1, totalRowsNeeded, canvasWidth)
    .setBorder(true, true, true, true, true, true, CONFIG.colors.graphLine, SpreadsheetApp.BorderStyle.DOTTED);

  // ==========================================
  // DRAW SIDE VIEW (Centered)
  // ==========================================
  let currentRow = startRow;
  
  planSheet.getRange(currentRow - 4, startCol).setValue("SIDE VIEW (Elevation)").setFontSize(12).setFontWeight("bold");

  heights_cells.forEach((hCells, index) => {
    let currentX = startCol;
    const hMeters = heights_meters[index];
    
    createLabelBox(planSheet, currentRow + Math.floor(hCells/2) - 1, startCol - 3, `${hMeters}m`, CONFIG.colors.dimText);

    for (let i = 0; i < spansX_cells.length; i++) {
      const wCells = spansX_cells[i];
      const room = planSheet.getRange(currentRow, currentX, hCells, wCells);
      room.setBackground(CONFIG.colors.fillSide);
      room.setBorder(true, true, true, true, null, null, CONFIG.colors.beam, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      currentX += wCells;
    }
    
    const floorNum = heights_cells.length - index;
    createLabelBox(planSheet, currentRow + Math.floor(hCells/2) - 1, currentX + 1, `FL ${floorNum}`, CONFIG.colors.gridLabel);

    currentRow += hCells;
  });

  // STUMP & SUPPORT
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

  // SIDE VIEW LABELS
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
  // DRAW TOP VIEW (Centered)
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