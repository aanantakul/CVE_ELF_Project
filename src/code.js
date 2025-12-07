/**
 * Building Structure Generator - V18 (Vertical Merge for Better Visibility)
 */

const CONFIG = {
  sheetPlan: "Plan",
  cellSizePx: 12,       
  resolution: 0.5,      // 1 Cell = 0.5m
  minPadding: 10,       
  stumpHeight: 2,
  pointLoadScale: 4,    // 1 ตัน = 4 ช่อง Grid
  colors: {
    beam: "#37474f",
    fillSide: "#f3e5f5",
    fillTop: "#e1f5fe",
    gridLabel: "#b71c1c",
    dimText: "#0d47a1",
    graphLine: "#eceff1",
    labelBg: "#ffffff",
    support: "#424242",
    loadArrow: "#b71c1c", 
    loadText: "#b71c1c"   
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏗️ CVE-RU ELF')
    .addItem('START ELF Program', 'showSidebar')
    .addToUi();
  showSidebar();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Building Generator')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html); 
}

function receiveFormInput(spanXStr, heightStr, loadsStr, spanYStr, pointLoadsStr) {
  generateBlueprintFromData(spanXStr, heightStr, loadsStr, spanYStr, pointLoadsStr);
}

// --- CORE LOGIC ---
function generateBlueprintFromData(rawSpanX, rawHeight, rawLoads, rawSpanY, rawPointLoads) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Default Values
  if (!rawSpanX) rawSpanX = "4,4,4";
  if (!rawHeight) rawHeight = "3.5,3.5";
  if (!rawLoads) rawLoads = "1.5, 2.0, 1.0"; 
  if (!rawPointLoads) rawPointLoads = "1.0, 2.5, 1.5"; 
  if (!rawSpanY) rawSpanY = "4,3";

  // 1. Parse Inputs
  const parseToCells = (str) => (str || "").toString().split(',').map(n => Math.round(parseFloat(n) / CONFIG.resolution));
  const parseToMeters = (str) => (str || "").toString().split(',').map(Number);
  const parseFloats = (str) => (str || "").toString().split(',').map(Number);

  const spansX_cells = parseToCells(rawSpanX);
  const heights_cells = parseToCells(rawHeight).reverse(); 
  const heights_meters = parseToMeters(rawHeight).reverse();
  const loads_val = parseFloats(rawLoads); 
  const pointLoads_val = parseFloats(rawPointLoads); 
  const spansY_cells = parseToCells(rawSpanY);
  const spansY_meters = parseToMeters(rawSpanY);

  // 2. คำนวณขนาด
  const drawingWidth = spansX_cells.reduce((a, b) => a + b, 0);
  const totalHeightCells_Side = heights_cells.reduce((a, b) => a + b, 0);
  const totalHeightCells_Top = spansY_cells.reduce((a, b) => a + b, 0);
  
  const maxPointLoad = Math.max(...pointLoads_val, 0);
  // เพิ่มพื้นที่ซ้ายสำหรับ Point Load (เผื่อ Merge 3 ช่อง + 2 ช่องเดิม)
  const requiredLeftSpace = Math.ceil(maxPointLoad * CONFIG.pointLoadScale) + 9; 
  
  const canvasWidth = Math.max(drawingWidth + requiredLeftSpace + CONFIG.minPadding, 80); 
  const totalRowsNeeded = totalHeightCells_Side + totalHeightCells_Top + 50; 

  let startCol = Math.floor((canvasWidth - drawingWidth) / 2);
  if (startCol < requiredLeftSpace) startCol = requiredLeftSpace; 
  
  let startRow = 8; 

  // 3. Setup Sheet
  const oldPlan = ss.getSheetByName(CONFIG.sheetPlan);
  if (oldPlan) oldPlan.setName("Plan_Old_Deleting");
  const planSheet = ss.insertSheet(CONFIG.sheetPlan);
  if (oldPlan) ss.deleteSheet(oldPlan);
  const inputSheet = ss.getSheetByName("Input");
  if (inputSheet) { try { ss.deleteSheet(inputSheet); } catch (e) {} }

  // 4. Resize
  if (canvasWidth > planSheet.getMaxColumns()) planSheet.insertColumnsAfter(planSheet.getMaxColumns(), canvasWidth - planSheet.getMaxColumns());
  if (planSheet.getMaxColumns() > canvasWidth) planSheet.deleteColumns(canvasWidth + 1, planSheet.getMaxColumns() - canvasWidth);
  if (totalRowsNeeded > planSheet.getMaxRows()) planSheet.insertRowsAfter(planSheet.getMaxRows(), totalRowsNeeded - planSheet.getMaxRows());

  planSheet.setColumnWidths(1, canvasWidth, CONFIG.cellSizePx);
  planSheet.setRowHeights(1, totalRowsNeeded, CONFIG.cellSizePx);
  planSheet.getRange(1, 1, totalRowsNeeded, canvasWidth).setBorder(true, true, true, true, true, true, CONFIG.colors.graphLine, SpreadsheetApp.BorderStyle.DOTTED);

  // ==========================================
  // DRAW SIDE VIEW
  // ==========================================
  let currentRow = startRow;
  
  planSheet.getRange(currentRow - 6, startCol).setValue("SIDE VIEW (Elevation)").setFontSize(12).setFontWeight("bold");

  // --- Loop วาดชั้นบนๆ ---
  heights_cells.forEach((hCells, index) => {
    let currentX = startCol;
    const hMeters = heights_meters[index];
    
    // Index Mapping (Bottom-Up Logic)
    // Point Load: Index 0=Base, 1=FL2, Last=Roof
    // Loop: Roof -> Base
    const levelIndex = (pointLoads_val.length - 1) - index; 
    
    // Dist Load: Index 0=BaseBeam, 1=FL2Beam, Last=RoofBeam
    const distLoadIndex = (loads_val.length - 1) - index;

    const distLoad = loads_val[distLoadIndex] || 0; 
    const pointLoad = pointLoads_val[levelIndex] || 0; 

    createLabelBox(planSheet, currentRow + Math.floor(hCells/2) - 1, startCol - 3, `${hMeters}m`, CONFIG.colors.dimText);
    
    // Point Load (Lateral)
    if (pointLoad > 0) {
      drawLateralLoad(planSheet, currentRow, startCol, pointLoad);
    }

    // Loop วาด Span
    for (let i = 0; i < spansX_cells.length; i++) {
      const wCells = spansX_cells[i];
      const room = planSheet.getRange(currentRow, currentX, hCells, wCells);
      room.setBackground(CONFIG.colors.fillSide);
      room.setBorder(true, true, true, true, null, null, CONFIG.colors.beam, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      if (distLoad > 0) {
        drawLoadArrows(planSheet, currentRow, currentX, wCells);
        if (i === 0) drawLoadLabel(planSheet, currentRow, currentX, wCells, distLoad);
      }
      currentX += wCells;
    }
    
    const floorNum = heights_cells.length - index;
    createLabelBox(planSheet, currentRow + Math.floor(hCells/2) - 1, currentX + 1, `FL ${floorNum}`, CONFIG.colors.gridLabel);

    currentRow += hCells;
  });

  // --- วาด Load ชั้นล่างสุด (Base / Ground Level) ---
  const basePointLoad = pointLoads_val[0] || 0;
  if (basePointLoad > 0) {
    drawLateralLoad(planSheet, currentRow, startCol, basePointLoad);
  }

  const bottomDistLoad = loads_val[0] || 0;
  if (bottomDistLoad > 0) {
    let currentX = startCol;
    for (let i = 0; i < spansX_cells.length; i++) {
      const wCells = spansX_cells[i];
      drawLoadArrows(planSheet, currentRow, currentX, wCells);
      if (i === 0) {
        drawLoadLabel(planSheet, currentRow, currentX, wCells, bottomDistLoad);
      }
      currentX += wCells;
    }
  }

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

  // LABELS & TOP VIEW
  let gridX = startCol;
  let gridNum = 1;
  const labelRow = currentRow + CONFIG.stumpHeight + 5; 

  createLabelBox(planSheet, labelRow, gridX - 1, gridNum++, CONFIG.colors.gridLabel);
  for (let i = 0; i < spansX_cells.length; i++) {
    const wCells = spansX_cells[i];
    const wMeters = (spansX_cells[i] * CONFIG.resolution).toFixed(1);
    createLabelBox(planSheet, labelRow, gridX + Math.floor(wCells/2) - 1, `${wMeters}m`, CONFIG.colors.dimText, "center", false);
    gridX += wCells;
    createLabelBox(planSheet, labelRow, gridX - 1, gridNum++, CONFIG.colors.gridLabel);
  }

  // Draw Top View
  currentRow += 15;
  planSheet.getRange(currentRow - 4, startCol).setValue("TOP VIEW (Plan)").setFontSize(12).setFontWeight("bold");

  gridX = startCol;
  gridNum = 1;
  createLabelBox(planSheet, currentRow - 3, gridX - 1, gridNum++, CONFIG.colors.gridLabel);

  for (let i = 0; i < spansX_cells.length; i++) {
    const wCells = spansX_cells[i];
    const wMeters = (spansX_cells[i] * CONFIG.resolution).toFixed(1);
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

// [แก้ไข] Point Load: ใช้ระดับเดิม (beamRow) แต่ Merge ขึ้นบน 2 ช่อง
function drawLateralLoad(sheet, beamRow, startCol, val) {
  const scale = CONFIG.pointLoadScale; 
  const arrowLength = Math.max(3, Math.ceil(val * scale));
  const extraSpace = 3; 
  const arrowStartCol = startCol - arrowLength - extraSpace;
  const totalWidth = arrowLength + extraSpace;

  if (arrowStartCol > 0) {
    // ใช้ beamRow - 1 เป็นจุดเริ่ม เพื่อ Merge ลงมาหา beamRow (รวม 2 ช่อง)
    // หรือจะเริ่มที่ beamRow-1 แล้วสูง 2 ก็ได้ เพื่อให้ครอบคลุมพื้นที่
    const targetRow = beamRow - 1; 
    const range = sheet.getRange(targetRow, arrowStartCol, 2, totalWidth); // สูง 2 ช่อง
    range.merge();
    
    const dashCount = Math.max(1, arrowLength); 
    const line = "─".repeat(dashCount); 
    const text = `${val}T ${line}→`;
    
    range.setValue(text);
    range.setHorizontalAlignment("right").setVerticalAlignment("middle");
    range.setFontColor(CONFIG.colors.loadArrow).setFontWeight("bold").setFontSize(9);
  }
}

function drawLoadArrows(sheet, beamRow, startCol, width) {
  const arrowRow = beamRow - 1;
  if (arrowRow > 0) {
    const arrowRange = sheet.getRange(arrowRow, startCol, 1, width);
    arrowRange.merge(); 
    const numArrows = Math.max(1, Math.floor(width - 1)); 
    const arrows = "↓ ↓ ".repeat(numArrows);
    arrowRange.setValue(arrows);
    arrowRange.setHorizontalAlignment("center").setVerticalAlignment("bottom");
    arrowRange.setFontColor(CONFIG.colors.loadArrow).setFontSize(8).setFontWeight("bold");
  }
}

// [แก้ไข] Dist Load Label: Merge ขึ้นบน 2 ช่อง
function drawLoadLabel(sheet, beamRow, startCol, width, val) {
  const textRow = beamRow - 2;
  if (textRow > 0) {
    // Merge แถว textRow-1 (บน) และ textRow (ล่าง) เข้าด้วยกัน
    const startRow = textRow - 1;
    const range = sheet.getRange(startRow, startCol, 2, width); // สูง 2 ช่อง
    range.merge();
    
    range.setValue(`${val} T/m`);
    range.setHorizontalAlignment("center").setVerticalAlignment("middle");
    range.setFontColor(CONFIG.colors.loadText).setFontSize(9).setFontWeight("bold");
  }
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
  const width = 4; const height = 2;
  const startX = centerX - Math.floor(width / 2);
  if (startX < 1) return;
  const range = sheet.getRange(row, startX, height, width);
  range.merge();
  range.setBackground(CONFIG.colors.support);
  range.setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  const soilRow = row + height; 
  const soilRange = sheet.getRange(soilRow, startX - 1, 1, width + 2); 
  soilRange.merge();
  soilRange.setValue("/ / / / / / / / / / / /"); 
  soilRange.setHorizontalAlignment("center").setVerticalAlignment("middle");
  soilRange.setFontSize(8).setFontColor("#757575"); 
  soilRange.setFontWeight("bold").setFontWeight("italic");
}

function drawColumnStump(sheet, row, x, height) {
  sheet.getRange(row, x, height, 1).setBorder(null, true, null, null, null, null, CONFIG.colors.beam, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}