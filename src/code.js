/**
 * Building Structure Generator & Seismic Analyzer - V37 (Fix Missing Shear Scale)
 */

const CONFIG = {
  sheetPlan: "Plan",
  sheetData: "Data_Seismic",
  sheetCalc: "Calculation_Report",
  cellSizePx: 12,       
  resolution: 0.5,      
  minPadding: 10,       
  stumpHeight: 2,
  pointLoadScale: 4,  // Scale แรงกระทำด้านข้าง (Fx): 1 ตัน = 4 ช่อง
  shearLoadScale: 2,  // Scale แรงเฉือน (V): 1 ตัน = 2 ช่อง
  colors: {
    beam: "#37474f", fillSide: "#f3e5f5", fillTop: "#e1f5fe",
    gridLabel: "#b71c1c", dimText: "#0d47a1", graphLine: "#eceff1",
    labelBg: "#ffffff", support: "#424242",
    loadArrow: "#b71c1c", loadText: "#b71c1c",
    shearArrow: "#00695c", shearText: "#004d40",
    tableHeader: "#1565c0", tableRowOdd: "#ffffff", tableRowEven: "#f5f5f5"
  }
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏗️ CVE-RU ELF').addItem('START ELF Program', 'showSidebar').addToUi();
  showSidebar();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('RC Seismic Analysis').setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html); 
}

function getSeismicDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.sheetData);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.sheetData);
    sheet.appendRow(["Province", "Amphoe", "Ss", "S1"]);
    sheet.appendRow(["SampleProv", "SampleDist", 0.5, 0.2]);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 4).getValues();
}

function receiveFormInput(spanXStr, heightStr, loadsStr, spanYStr, seismicParams) {
  let calcResult = null; 
  let pointLoadsStr = "0"; 

  if (seismicParams && seismicParams.ss !== "-" && seismicParams.ss !== "") {
    calcResult = calculateELF(spanXStr, heightStr, loadsStr, seismicParams);
    pointLoadsStr = calcResult.pointLoadsStr;
  }
  
  generateBlueprintFromData(spanXStr, heightStr, loadsStr, spanYStr, pointLoadsStr, calcResult);

  if (calcResult) {
    createCalculationSheet(calcResult);
  }
}

// --- CALCULATION LOGIC ---
function calculateELF(spanXStr, heightStr, loadsStr, params) {
  const spanX = (spanXStr || "").toString().split(',').map(Number);
  const heights = (heightStr || "").toString().split(',').map(Number); 
  const loads = (loadsStr || "").toString().split(',').map(Number);    
  
  const Ss = parseFloat(params.ss); const S1 = parseFloat(params.s1);
  const SiteClass = params.siteClass;
  const Ie = parseFloat(params.Ie); const R = parseFloat(params.R);

  const Fa = getFa(SiteClass, Ss);
  const Fv = getFv(SiteClass, S1);
  const SMS = Fa * Ss; const SM1 = Fv * S1;
  const SDS = (2/3) * SMS; const SD1 = (2/3) * SM1;

  const totalHeight = heights.reduce((a,b) => a+b, 0);
  const T = 0.0466 * Math.pow(totalHeight, 0.9); 

  let Cs = SDS / (R / Ie);
  const Cs_max = SD1 / (T * (R / Ie));
  if (Cs > Cs_max) Cs = Cs_max;
  let Cs_min = 0.01;
  if (Cs < Cs_min) Cs = Cs_min;

  const totalSpanLength = spanX.reduce((a,b) => a+b, 0);
  let levels = []; 
  let currentHeight = 0; 

  levels.push({ name: "Base/FL1", hx: 0, wx: loads[0] * totalSpanLength, wx_hx_k: 0, Fx: 0, Vx: 0 });

  for (let i = 0; i < heights.length; i++) {
    currentHeight += heights[i];
    const loadIndex = i + 1;
    const w_i = (loads[loadIndex] || 0) * totalSpanLength;
    const name = (i === heights.length - 1) ? "Roof" : `FL ${i+2}`;
    levels.push({ name: name, hx: currentHeight, wx: w_i, wx_hx_k: 0 });
  }

  let W_effective = 0;
  levels.forEach(l => { if (l.hx > 0) W_effective += l.wx; });
  const V = Cs * W_effective;

  let k = 1;
  if (T <= 0.5) k = 1; else if (T >= 2.5) k = 2; else k = 1 + ((T - 0.5) / 2);

  let sum_w_h_k = 0;
  levels.forEach(l => {
    if (l.hx > 0) { l.wx_hx_k = l.wx * Math.pow(l.hx, k); sum_w_h_k += l.wx_hx_k; }
  });

  let pointLoadsArr = [];
  levels.forEach(l => {
    if (l.hx > 0 && sum_w_h_k > 0) {
      const Cvx = l.wx_hx_k / sum_w_h_k;
      l.Fx = Cvx * V;
    } else { l.Fx = 0; }
    pointLoadsArr.push(l.Fx.toFixed(3));
  });

  let cumShear = 0;
  for (let i = levels.length - 1; i >= 0; i--) {
    cumShear += levels[i].Fx;
    levels[i].Vx = cumShear;
  }

  return { pointLoadsStr: pointLoadsArr.join(","), levels: levels, params: { Fa, Fv, SDS, SD1, T, Cs, V, W_effective, k, Ss, S1, R, Ie, SiteClass } };
}

// --- INTERPOLATION ---
function getFa(siteClass, Ss) {
  const grid = { "A": [0.8,0.8,0.8,0.8,0.8], "B": [1.0,1.0,1.0,1.0,1.0], "C": [1.2,1.2,1.1,1.0,1.0], "D": [1.6,1.4,1.2,1.1,1.0], "E": [2.5,1.7,1.2,0.9,0.9], "F": [1,1,1,1,1] };
  return interpolate(Ss, [0.25, 0.50, 0.75, 1.00, 1.25], grid[siteClass] || grid["D"]);
}
function getFv(siteClass, S1) {
  const grid = { "A": [0.8,0.8,0.8,0.8,0.8], "B": [1.0,1.0,1.0,1.0,1.0], "C": [1.7,1.6,1.5,1.4,1.3], "D": [2.4,2.0,1.8,1.6,1.5], "E": [3.5,3.2,2.8,2.4,2.4], "F": [1,1,1,1,1] };
  return interpolate(S1, [0.1, 0.2, 0.3, 0.4, 0.5], grid[siteClass] || grid["D"]);
}
function interpolate(x, x_arr, y_arr) {
  if (x <= x_arr[0]) return y_arr[0];
  if (x >= x_arr[x_arr.length - 1]) return y_arr[y_arr.length - 1];
  for (let i = 0; i < x_arr.length - 1; i++) {
    if (x >= x_arr[i] && x <= x_arr[i+1]) {
      const x_lower = x_arr[i]; const x_upper = x_arr[i+1];
      const y_lower = y_arr[i]; const y_upper = y_arr[i+1];
      return y_lower + (x - x_lower) * (y_upper - y_lower) / (x_upper - x_lower);
    }
  }
  return y_arr[0];
}

// --- CALCULATION REPORT ---
function createCalculationSheet(result) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.sheetCalc);
  if (!sheet) { sheet = ss.insertSheet(CONFIG.sheetCalc); } else { sheet.clear(); }

  const p = result.params;
  const levels = result.levels.slice().reverse(); 

  sheet.setColumnWidth(1, 80); sheet.setColumnWidth(2, 80); sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 120); sheet.setColumnWidth(5, 80); sheet.setColumnWidth(6, 100); sheet.setColumnWidth(7, 100);

  let r = 1;
  sheet.getRange(r, 1).setValue("SEISMIC ANALYSIS REPORT (DPT 1301/1302-61)").setFontSize(14).setFontWeight("bold");
  r += 2;

  const paramsData = [
    ["Parameter", "Value", "Description"],
    ["Ss", p.Ss, "Spectral Acceleration (Short Period)"],
    ["S1", p.S1, "Spectral Acceleration (1.0s)"],
    ["Site Class", p.SiteClass, "Soil Type"],
    ["Fa", p.Fa.toFixed(2), "Site Coefficient (Short)"],
    ["Fv", p.Fv.toFixed(2), "Site Coefficient (Long)"],
    ["SMS", (p.Fa * p.Ss).toFixed(3), "Adjusted Spectral Acc. (Short)"],
    ["SM1", (p.Fv * p.S1).toFixed(3), "Adjusted Spectral Acc. (1.0s)"],
    ["SDS", p.SDS.toFixed(3), "Design Spectral Acc. (Short)"],
    ["SD1", p.SD1.toFixed(3), "Design Spectral Acc. (1.0s)"],
    ["R", p.R, "Response Modification Coefficient"],
    ["Ie", p.Ie, "Importance Factor"],
    ["T (Period)", p.T.toFixed(3) + " s", "Approximate Fundamental Period"],
    ["k", p.k.toFixed(2), "Distribution Exponent"],
    ["W (Weight)", p.W_effective.toFixed(2) + " T", "Total Seismic Weight"],
    ["Cs", p.Cs.toFixed(4), "Seismic Response Coefficient"],
    ["V (Base Shear)", p.V.toFixed(2) + " T", "Design Base Shear (V = Cs * W)"]
  ];

  const paramRange = sheet.getRange(r, 1, paramsData.length, 3);
  paramRange.setValues(paramsData).setBorder(true, true, true, true, true, true, "#999999", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(r, 1, 1, 3).setBackground(CONFIG.colors.tableHeader).setFontColor("white").setFontWeight("bold");
  sheet.getRange(r, 2, paramsData.length, 1).setHorizontalAlignment("left"); // Left align Value column

  r += paramsData.length + 2;

  sheet.getRange(r, 1).setValue("VERTICAL DISTRIBUTION OF SEISMIC FORCES").setFontWeight("bold");
  r++;

  const headerRow = ["Level", "Height hx (m)", "Weight wx (T)", "wx * hx^k", "Cvx", "Force Fx (T)", "Shear Vx (T)"];
  sheet.getRange(r, 1, 1, 7).setValues([headerRow]).setBackground(CONFIG.colors.tableHeader).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  r++;

  let tableData = [];
  levels.forEach(l => {
    const Cvx = (l.Fx / p.V) || 0;
    tableData.push([l.name, l.hx.toFixed(2), l.wx.toFixed(2), l.wx_hx_k.toFixed(1), Cvx.toFixed(4), l.Fx.toFixed(3), l.Vx ? l.Vx.toFixed(3) : "-"]);
  });

  const dataRange = sheet.getRange(r, 1, tableData.length, 7);
  dataRange.setValues(tableData).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment("center");
  for (let i = 0; i < tableData.length; i++) { if (i % 2 !== 0) sheet.getRange(r + i, 1, 1, 7).setBackground(CONFIG.colors.tableRowEven); }
  sheet.autoResizeColumns(1, 7);
}

// ==========================================
// 🎨 DRAWING LOGIC (Blueprint)
// ==========================================
function generateBlueprintFromData(rawSpanX, rawHeight, rawLoads, rawSpanY, rawPointLoads, calcResult) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let planSheet = ss.getSheetByName(CONFIG.sheetPlan);
  if (!planSheet) { planSheet = ss.insertSheet(CONFIG.sheetPlan); } else { planSheet.clear(); }

  if (!rawSpanX) rawSpanX = "4,4,4"; if (!rawHeight) rawHeight = "3.5,3.5";
  if (!rawLoads) rawLoads = "1.5, 2.0, 1.0"; if (!rawPointLoads) rawPointLoads = "0,0,0"; 
  if (!rawSpanY) rawSpanY = "4,3";

  const parseToCells = (str) => (str || "").toString().split(',').map(n => Math.round(parseFloat(n) / CONFIG.resolution));
  const parseFloats = (str) => (str || "").toString().split(',').map(Number);

  const spansX_cells = parseToCells(rawSpanX);
  const heights_cells = parseToCells(rawHeight).reverse(); 
  const heights_meters = parseFloats(rawHeight).reverse();
  const loads_val = parseFloats(rawLoads); 
  const pointLoads_val = parseFloats(rawPointLoads); 
  const spansY_cells = parseToCells(rawSpanY);

  const drawingWidth = spansX_cells.reduce((a, b) => a + b, 0);
  const totalHeightCells = heights_cells.reduce((a, b) => a + b, 0) + spansY_cells.reduce((a, b) => a + b, 0);
  
  const maxPointLoad = Math.max(...pointLoads_val, 0);
  const requiredLeftSpace = Math.ceil(maxPointLoad * CONFIG.pointLoadScale) + 15; 
  
  // คำนวณพื้นที่ขวาสำหรับ Shear V (ใช้ shearLoadScale)
  let maxShear = 0;
  if (calcResult && calcResult.params && calcResult.params.V) maxShear = calcResult.params.V;
  const requiredRightSpace = Math.ceil(maxShear * CONFIG.shearLoadScale) + 15; 

  const canvasWidth = Math.max(drawingWidth + requiredLeftSpace + requiredRightSpace + CONFIG.minPadding, 80); 
  const totalRowsNeeded = totalHeightCells + 50; 

  if (canvasWidth > planSheet.getMaxColumns()) planSheet.insertColumnsAfter(planSheet.getMaxColumns(), canvasWidth - planSheet.getMaxColumns());
  if (totalRowsNeeded > planSheet.getMaxRows()) planSheet.insertRowsAfter(planSheet.getMaxRows(), totalRowsNeeded - planSheet.getMaxRows());
  
  planSheet.setColumnWidths(1, canvasWidth, CONFIG.cellSizePx);
  planSheet.setRowHeights(1, totalRowsNeeded, CONFIG.cellSizePx);
  
  planSheet.getRange(1, 1, totalRowsNeeded, canvasWidth).setBorder(true, true, true, true, true, true, CONFIG.colors.graphLine, SpreadsheetApp.BorderStyle.DOTTED);

  let startCol = Math.floor((canvasWidth - drawingWidth) / 2);
  if (startCol < requiredLeftSpace) startCol = requiredLeftSpace; 
  let startRow = 8; 

  // --- DRAW SIDE VIEW ---
  let currentRow = startRow;
  planSheet.getRange(currentRow - 6, startCol).setValue("SIDE VIEW (Elevation)").setFontSize(12).setFontWeight("bold");

  heights_cells.forEach((hCells, index) => {
    let currentX = startCol;
    const hMeters = heights_meters[index];
    const levelIndex = (pointLoads_val.length - 1) - index; 
    const distLoadIndex = (loads_val.length - 1) - index;
    const distLoad = loads_val[distLoadIndex] || 0; 
    const pointLoad = pointLoads_val[levelIndex] || 0; 

    createLabelBox(planSheet, currentRow + Math.floor(hCells/2) - 1, startCol - 3, `${hMeters}m`, CONFIG.colors.dimText);
    
    // Draw Left Point Load (Fx)
    if (pointLoad > 0) drawLateralLoad(planSheet, currentRow, startCol, pointLoad);

    // Draw Right Shear Force (Vx)
    if (calcResult && calcResult.levels) {
       const shearIndex = (calcResult.levels.length - 1) - index;
       const shearVal = calcResult.levels[shearIndex] ? calcResult.levels[shearIndex].Vx : 0;
       
       if (shearVal > 0) {
         const rightEdgeCol = startCol + drawingWidth;
         drawShearForceRight(planSheet, currentRow, rightEdgeCol, hCells, shearVal);
       }
    }

    for (let i = 0; i < spansX_cells.length; i++) {
      const wCells = spansX_cells[i];
      const room = planSheet.getRange(currentRow, currentX, hCells, wCells);
      room.setBackground(CONFIG.colors.fillSide).setBorder(true, true, true, true, null, null, CONFIG.colors.beam, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      if (distLoad > 0) {
        drawLoadArrows(planSheet, currentRow, currentX, wCells);
        if (i === 0) drawLoadLabel(planSheet, currentRow, currentX, wCells, distLoad);
      }
      currentX += wCells;
    }
    
    const floorNum = heights_cells.length - index;
    createLabelBox(planSheet, currentRow + 1, currentX - 3, `FL ${floorNum}`, CONFIG.colors.gridLabel);

    currentRow += hCells;
  });

  const basePointLoad = pointLoads_val[0] || 0;
  if (basePointLoad > 0) drawLateralLoad(planSheet, currentRow, startCol, basePointLoad);

  const bottomDistLoad = loads_val[0] || 0;
  if (bottomDistLoad > 0) {
    let currentX = startCol;
    for (let i = 0; i < spansX_cells.length; i++) {
      const wCells = spansX_cells[i];
      drawLoadArrows(planSheet, currentRow, currentX, wCells);
      if (i === 0) drawLoadLabel(planSheet, currentRow, currentX, wCells, bottomDistLoad);
      currentX += wCells;
    }
  }

  let columnX = startCol;
  const colPositions = [startCol];
  spansX_cells.forEach(w => { columnX += w; colPositions.push(columnX); });
  colPositions.forEach(x => {
    drawColumnStump(planSheet, currentRow, x, CONFIG.stumpHeight);
    drawFixedSupport(planSheet, currentRow + CONFIG.stumpHeight, x);
  });

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

  currentRow += 15;
  planSheet.getRange(currentRow - 4, startCol).setValue("TOP VIEW (Plan)").setFontSize(12).setFontWeight("bold");
  gridX = startCol; gridNum = 1;
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
    const charCode = 65 + index;
    createLabelBox(planSheet, currentRow - 1, startCol - 3, String.fromCharCode(charCode), CONFIG.colors.gridLabel);
    createLabelBox(planSheet, currentRow + Math.floor(hCells/2) - 1, startCol - 3, `${(spansY_cells[index] * CONFIG.resolution).toFixed(1)}m`, CONFIG.colors.dimText, "center", false);
    for (let i = 0; i < spansX_cells.length; i++) {
      const wCells = spansX_cells[i];
      const room = planSheet.getRange(currentRow, currentX, hCells, wCells);
      room.setBackground(CONFIG.colors.fillTop).setBorder(true, true, true, true, null, null, "#90a4ae", SpreadsheetApp.BorderStyle.SOLID);
      currentX += wCells;
    }
    if (index === spansY_cells.length - 1) {
       createLabelBox(planSheet, currentRow + hCells - 1, startCol - 3, String.fromCharCode(charCode + 1), CONFIG.colors.gridLabel);
    }
    currentRow += hCells;
  });
  
  planSheet.setHiddenGridlines(true);
}

// --- HELPERS ---

// [แก้ไข] ใช้ shearLoadScale
function drawShearForceRight(sheet, row, col, height, val) {
  const scale = CONFIG.shearLoadScale; // 0.5
  const arrowLength = Math.max(3, Math.ceil(val * scale));
  const dashCount = Math.max(1, arrowLength); 
  const line = "─".repeat(dashCount); 
  const text = `←${line} V = ${val.toFixed(2)}T`;

  const arrowRow = row + Math.floor(height / 2);
  const totalWidth = arrowLength + 10; 

  // วาดที่ตำแหน่ง col (ชิดผนัง)
  const range = sheet.getRange(arrowRow, col, 1, totalWidth);
  range.merge();
  range.setValue(text);
  range.setHorizontalAlignment("left").setVerticalAlignment("middle");
  range.setFontColor(CONFIG.colors.shearArrow).setFontWeight("bold").setFontSize(9);
}

function drawLateralLoad(sheet, beamRow, startCol, val) {
  const scale = CONFIG.pointLoadScale; 
  const arrowLength = Math.max(3, Math.ceil(val * scale));
  const extraSpace = 6; 
  const arrowStartCol = startCol - arrowLength - extraSpace;
  const totalWidth = arrowLength + extraSpace;
  if (arrowStartCol > 0) {
    const targetRow = beamRow - 1; 
    const range = sheet.getRange(targetRow, arrowStartCol, 2, totalWidth); 
    range.merge();
    const dashCount = Math.max(1, arrowLength); 
    const line = "─".repeat(dashCount); 
    const valStr = (typeof val === 'number') ? val.toFixed(3) : val;
    const text = `F = ${valStr}T ${line}→`;
    range.setValue(text).setHorizontalAlignment("right").setVerticalAlignment("middle").setFontColor(CONFIG.colors.loadArrow).setFontWeight("bold").setFontSize(9);
  }
}
function drawLoadArrows(sheet, beamRow, startCol, width) {
  const arrowRow = beamRow - 1;
  if (arrowRow > 0) {
    const arrowRange = sheet.getRange(arrowRow, startCol, 1, width);
    arrowRange.merge().setValue("↓ ↓ ".repeat(Math.max(1, Math.floor(width - 1))))
      .setHorizontalAlignment("center").setVerticalAlignment("bottom").setFontColor(CONFIG.colors.loadArrow).setFontSize(8).setFontWeight("bold");
  }
}
function drawLoadLabel(sheet, beamRow, startCol, width, val) {
  const textRow = beamRow - 2;
  if (textRow > 0) {
    const range = sheet.getRange(textRow - 1, startCol, 2, width); 
    range.merge().setValue(`${val} T/m`).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontColor(CONFIG.colors.loadText).setFontSize(9).setFontWeight("bold");
  }
}
function createLabelBox(sheet, row, col, text, color, align = "center", isBold = true) {
  if (row < 1 || col < 1) return; 
  const range = sheet.getRange(row, col, 2, 2); 
  range.merge().setValue(text).setFontColor(color).setBackground(CONFIG.colors.labelBg).setBorder(true, true, true, true, null, null, "#dddddd", SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment(align).setVerticalAlignment("middle").setFontSize(8);
  if (isBold) range.setFontWeight("bold");
}
function drawFixedSupport(sheet, row, centerX) {
  const width = 4; const height = 2; const startX = centerX - Math.floor(width / 2);
  if (startX < 1) return;
  const range = sheet.getRange(row, startX, height, width);
  range.merge().setBackground(CONFIG.colors.support).setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  const soilRow = row + height; 
  const soilRange = sheet.getRange(soilRow, startX - 1, 1, width + 2); 
  soilRange.merge().setValue("/ / / / / / / / / / / /").setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(8).setFontColor("#757575").setFontWeight("bold").setFontWeight("italic");
}
function drawColumnStump(sheet, row, x, height) {
  sheet.getRange(row, x, height, 1).setBorder(null, true, null, null, null, null, CONFIG.colors.beam, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}