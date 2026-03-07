/**
 * Building Structure Generator & Seismic Analyzer - V0.5.6
 * + FIXED: Merged Cell Overlap Exception (Dimension vs Force Arrows)
 * + FORMAT: Left-Aligned Drawing, Clean Section Texts, Full-width Spans
 */

const CONFIG = {
  sheetPlan: "Plan",
  sheetData: "Data_Seismic",
  sheetCalc: "Calculation_Report",
  cellSizePx: 12,       
  defaultResolution: 0.5, 
  minPadding: 4,        
  stumpHeight: 2,
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
    loadText: "#b71c1c",
    shearArrow: "#00695c", 
    shearText: "#004d40",
    tableHeader: "#1565c0", 
    tableRowEven: "#f5f5f5",
    warningText: "#d93025"
  }
};

const BANGKOK_ZONE_DATA = {
  1:  { Sa05: 0.360 }, 2:  { Sa05: 0.352 }, 3:  { Sa05: 0.262 }, 4:  { Sa05: 0.287 }, 5:  { Sa05: 0.191 }, 
  6:  { Sa05: 0.272 }, 7:  { Sa05: 0.246 }, 8:  { Sa05: 0.162 }, 9:  { Sa05: 0.214 }, 10: { Sa05: 0.179 }
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏗️ CVE-RU ELF')
    .addItem('▶️ START ELF Program', 'showSidebar')
    .addSeparator()
    .addItem('🔄 Reset/Load Database', 'setupSeismicDatabase')
    .addToUi();
  showSidebar();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('RC Seismic Analysis')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html); 
}

function getSeismicDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.sheetData);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.sheetData);
    sheet.appendRow(["Province", "Amphoe", "Ss", "S1", "Zone (0=Out, 1-10=Basin)"]);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 5).getValues();
}

function setupSeismicDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.sheetData);
  
  if (sheet) { sheet.clear(); } else { sheet = ss.insertSheet(CONFIG.sheetData); }
  
  const headers = [["Province", "Amphoe", "Ss", "S1", "Zone"]];
  const data = [
    ["กรุงเทพมหานคร", "พระนคร", 0, 0, 1], ["กรุงเทพมหานคร", "จตุจักร", 0, 0, 2],
    ["นนทบุรี", "เมืองนนทบุรี", 0, 0, 3], ["สมุทรปราการ", "เมืองสมุทรปราการ", 0, 0, 7],
    ["เชียงใหม่", "เมืองเชียงใหม่", 0.963, 0.248, 0], ["กาญจนบุรี", "เมืองกาญจนบุรี", 0.642, 0.241, 0]
  ];
  
  sheet.getRange(1, 1, 1, 5).setValues(headers).setFontWeight("bold").setBackground("#e6f4ea");
  sheet.getRange(2, 1, data.length, 5).setValues(data);
  sheet.setColumnWidth(1, 120); sheet.setColumnWidth(2, 120); sheet.setColumnWidth(5, 80);
  
  SpreadsheetApp.getUi().alert('✅ Database Initialized!');
}

function receiveFormInput(spanXStr, heightStr, loadsStr, spanYStr, seismicParams) {
  let calcResult = null; 
  let pointLoadsStr = "0"; 
  
  if (seismicParams) {
    calcResult = calculateELF(spanXStr, heightStr, loadsStr, spanYStr, seismicParams);
    calcResult = calculateStability(calcResult, seismicParams, spanXStr, spanYStr, heightStr);
    pointLoadsStr = calcResult.pointLoadsStr;
  }
  
  if (calcResult) { createCalculationSheet(calcResult); }
  generateBlueprintFromData(spanXStr, heightStr, loadsStr, spanYStr, pointLoadsStr, calcResult);
}

// --- MODULE 1: ELF CALCULATION ---
function calculateELF(spanXStr, heightStr, loadsStr, spanYStr, params) {
  const spanX = (spanXStr || "").toString().split(',').map(Number);
  const spanY = (spanYStr || "").toString().split(',').map(Number);
  const heights = (heightStr || "").toString().split(',').map(Number); 
  const loads = (loadsStr || "").toString().split(',').map(Number);    
  
  const Ss = parseFloat(params.ss) || 0; 
  const S1 = parseFloat(params.s1) || 0;
  const SiteClass = params.siteClass;
  const RiskCat = params.riskCat;
  const Ie = parseFloat(params.Ie) || 1.0; 
  const R = parseFloat(params.R) || 3.0;
  
  let totalHeight = 0;
  for (let i = 0; i < heights.length; i++) totalHeight += (heights[i] || 0);
  const T = 0.02 * totalHeight; 

  let SDS = 0, SD1 = 0, Sa05 = 0, Fa = 0, Fv = 0;
  let designCategory = "";
  let calculationMethod = "";
  let Cs = 0;

  if (params.basinZone > 0) {
    calculationMethod = "Bangkok Basin (Zone " + params.basinZone + ") - Using Sa(0.5s)";
    const zoneData = BANGKOK_ZONE_DATA[params.basinZone];
    Sa05 = zoneData ? zoneData.Sa05 : 0;
    
    if (T <= 0.5) {
      designCategory = classifySDC_Table223(Sa05, RiskCat);
      calculationMethod += " | T=" + T.toFixed(2) + "s (<=0.5s) -> Use Table 2.2.3";
    } else {
      designCategory = classifySDC_Table224(Sa05, RiskCat);
      calculationMethod += " | T=" + T.toFixed(2) + "s (>0.5s) -> Use Table 2.2.4";
    }
    SDS = Sa05; SD1 = 0; Cs = Sa05 / (R / Ie); 
  } else {
    calculationMethod = "Standard ELF (DPT 1301/1302)";
    Fa = getFa(SiteClass, Ss); Fv = getFv(SiteClass, S1);
    SDS = (2/3) * Fa * Ss; SD1 = (2/3) * Fv * S1;
    const catShort = classifySDC_Table223(SDS, RiskCat);
    const catLong = classifySDC_Table224(SD1, RiskCat);
    designCategory = (catShort > catLong) ? catShort : catLong;
    Cs = SDS / (R / Ie);
    const Cs_max = SD1 / (T * (R / Ie));
    if (Cs > Cs_max) Cs = Cs_max;
    if (Ss === 0 && S1 === 0) Cs = 0; else if (Cs < 0.01) Cs = 0.01;
  }

  let totalSpanLength = 0; for (let i = 0; i < spanX.length; i++) totalSpanLength += (spanX[i] || 0);
  let totalDepth = 0; for (let i = 0; i < spanY.length; i++) totalDepth += (spanY[i] || 0);

  let levels = []; 
  let currentHeight = 0; 
  levels.push({ name: "Base/FL1", hx: 0, wx: (loads[0] || 0) * totalSpanLength * totalDepth, wx_hx_k: 0, Fx: 0, Vx: 0 });

  for (let i = 0; i < heights.length; i++) {
    currentHeight += (heights[i] || 0);
    const loadIndex = i + 1;
    const isRoof = (i === heights.length - 1);
    levels.push({ name: isRoof ? "Roof" : "FL " + (i + 2), hx: currentHeight, wx: (loads[loadIndex] || 0) * totalSpanLength * totalDepth, wx_hx_k: 0 });
  }

  let W_effective = 0;
  for (let i = 0; i < levels.length; i++) if (levels[i].hx > 0) W_effective += levels[i].wx;
  
  const V = Cs * W_effective;
  let k = 1; if (T <= 0.5) k = 1; else if (T >= 2.5) k = 2; else k = 1 + ((T - 0.5) / 2);

  let sum_w_h_k = 0;
  for (let i = 0; i < levels.length; i++) {
    if (levels[i].hx > 0) {
      levels[i].wx_hx_k = levels[i].wx * Math.pow(levels[i].hx, k);
      sum_w_h_k += levels[i].wx_hx_k;
    }
  }

  let pointLoadsArr = [];
  for (let i = 0; i < levels.length; i++) {
    if (levels[i].hx > 0 && sum_w_h_k > 0) levels[i].Fx = (levels[i].wx_hx_k / sum_w_h_k) * V;
    else levels[i].Fx = 0;
    pointLoadsArr.push((levels[i].Fx || 0).toFixed(3));
  }

  let cumShear = 0;
  for (let i = levels.length - 1; i >= 0; i--) {
    cumShear += levels[i].Fx;
    levels[i].Vx = cumShear;
  }

  return { 
    pointLoadsStr: pointLoadsArr.join(","), levels: levels, 
    params: { 
      Sa05: Sa05, T: T, Cs: Cs, V: V, Weff: W_effective, Ss: Ss, S1: S1, SDS: SDS, SD1: SD1, 
      Fa: Fa, Fv: Fv, SiteClass: SiteClass, RiskCat: RiskCat, R: R, Ie: Ie, dc: designCategory, meth: calculationMethod, 
      basinZone: params.basinZone, systemType: params.systemType, isAllowed: params.isAllowed, violationMsg: params.violationMsg,
      stiffnessConfig: params.stiffnessConfig
    } 
  };
}

// --- MODULE 2: DUAL-AXIS STABILITY & DRIFT ---
function calculateStability(calcResult, params, spanXStr, spanYStr, heightsStr) {
  const levels = calcResult.levels; const sParams = params.stiffnessConfig;
  if (!sParams) return calcResult; 

  const spanX = (spanXStr || "").toString().split(',').map(Number);
  const spanY = (spanYStr || "").toString().split(',').map(Number);
  const heights = (heightsStr || "").toString().split(',').map(Number);
  
  const nx = spanX.length; const ny = spanY.length; const numCols = (nx + 1) * (ny + 1);
  let sum_1_Lx = 0; for (let i = 0; i < spanX.length; i++) if(spanX[i]) sum_1_Lx += (1 / spanX[i]);
  let sum_1_Ly = 0; for (let i = 0; i < spanY.length; i++) if(spanY[i]) sum_1_Ly += (1 / spanY[i]);
  
  const frameFactor_Ib_Lx = (ny + 1) * sum_1_Lx; const frameFactor_Ib_Ly = (nx + 1) * sum_1_Ly; 
  const Ec = 15100 * Math.sqrt(sParams.fc || 280) * 10; 
  const Cd = parseFloat(params.Cd) || 2.5; const Ie = parseFloat(params.Ie) || 1.0;
  
  for (let i = 1; i < levels.length; i++) {
    let hs = heights[i-1] || 3.5; let Vx = levels[i].Vx || 0;   
    let Px = 0; for (let j = i; j < levels.length; j++) Px += levels[j].wx;
    
    let col_b = sParams.col_b || 0.4; let col_h = sParams.col_h || 0.4;
    let bxb = sParams.bmX_b || 0.2; let bxh = sParams.bmX_h || 0.4;
    let byb = sParams.bmY_b || 0.2; let byh = sParams.bmY_h || 0.4;
    
    let Ig_cy = (col_h * Math.pow(col_b, 3)) / 12; let Ig_bx = (bxb * Math.pow(bxh, 3)) / 12;
    let stiffness_col_X = numCols * ((0.70 * Ig_cy) / hs); let stiffness_bm_X = frameFactor_Ib_Lx * (0.35 * Ig_bx);
    let kx_X = (12 * Ec) / (hs * hs * ((1 / stiffness_col_X) + (1 / stiffness_bm_X)));
    let delta_xe_X = (kx_X > 0) ? (Vx / kx_X) : 0;
    let delta_x_X = delta_xe_X * (Cd / Ie); let driftRatio_X = delta_x_X / hs;
    
    let Ig_cx = (col_b * Math.pow(col_h, 3)) / 12; let Ig_by = (byb * Math.pow(byh, 3)) / 12;
    let stiffness_col_Y = numCols * ((0.70 * Ig_cx) / hs); let stiffness_bm_Y = frameFactor_Ib_Ly * (0.35 * Ig_by);
    let kx_Y = (12 * Ec) / (hs * hs * ((1 / stiffness_col_Y) + (1 / stiffness_bm_Y)));
    let delta_xe_Y = (kx_Y > 0) ? (Vx / kx_Y) : 0;
    let delta_x_Y = delta_xe_Y * (Cd / Ie); let driftRatio_Y = delta_x_Y / hs;
    
    let theta_max = 0.5 / Cd; if (theta_max > 0.25) theta_max = 0.25;
    let theta_X = 0; let pStatus_X = "OK";
    if (Vx > 0 && hs > 0 && Cd > 0) { theta_X = (Px * delta_x_X) / (Vx * hs * Cd); if (theta_X > theta_max) pStatus_X = "FAIL"; else if (theta_X > 0.1) pStatus_X = "Amplify"; }

    let theta_Y = 0; let pStatus_Y = "OK";
    if (Vx > 0 && hs > 0 && Cd > 0) { theta_Y = (Px * delta_x_Y) / (Vx * hs * Cd); if (theta_Y > theta_max) pStatus_Y = "FAIL"; else if (theta_Y > 0.1) pStatus_Y = "Amplify"; }

    let Mx = 0; let h_base = levels[i-1].hx; for (let j = i; j < levels.length; j++) Mx += levels[j].Fx * (levels[j].hx - h_base);
    
    levels[i].Px = Px; levels[i].Mx = Mx; levels[i].theta_max = theta_max;
    levels[i].kx_X = kx_X; levels[i].dx_X = delta_x_X * 1000; levels[i].drift_X = driftRatio_X; levels[i].dStat_X = (driftRatio_X <= 0.01) ? "OK" : "FAIL"; levels[i].th_X = theta_X; levels[i].pStat_X = pStatus_X;
    levels[i].kx_Y = kx_Y; levels[i].dx_Y = delta_x_Y * 1000; levels[i].drift_Y = driftRatio_Y; levels[i].dStat_Y = (driftRatio_Y <= 0.01) ? "OK" : "FAIL"; levels[i].th_Y = theta_Y; levels[i].pStat_Y = pStatus_Y;
  }
  return calcResult;
}

// --- HELPER FUNCTIONS ---
function classifySDC_Table223(val, riskCat) { let rLvl = 0; if (riskCat === 'III') rLvl = 1; if (riskCat === 'IV') rLvl = 2; const cats = ["A", "B", "C", "D"]; if (val < 0.167) return cats[0]; if (val < 0.33) return cats[(rLvl === 2) ? 2 : 1]; if (val < 0.50) return cats[(rLvl === 2) ? 3 : 2]; return cats[3]; }
function classifySDC_Table224(val, riskCat) { let rLvl = 0; if (riskCat === 'III') rLvl = 1; if (riskCat === 'IV') rLvl = 2; const cats = ["A", "B", "C", "D"]; if (val < 0.067) return cats[0]; if (val < 0.133) return cats[(rLvl === 2) ? 2 : 1]; if (val < 0.20) return cats[(rLvl === 2) ? 3 : 2]; return cats[3]; }
function getFa(siteClass, Ss) { const grid = { "A": [0.8, 0.8, 0.8, 0.8, 0.8], "B": [1.0, 1.0, 1.0, 1.0, 1.0], "C": [1.2, 1.2, 1.1, 1.0, 1.0], "D": [1.6, 1.4, 1.2, 1.1, 1.0], "E": [2.5, 1.7, 1.2, 0.9, 0.9] }; const y_arr = grid[siteClass] || grid["D"]; return interpolateLine(Ss, [0.25, 0.5, 0.75, 1.0, 1.25], y_arr); }
function getFv(siteClass, S1) { const grid = { "A": [0.8, 0.8, 0.8, 0.8, 0.8], "B": [1.0, 1.0, 1.0, 1.0, 1.0], "C": [1.7, 1.6, 1.5, 1.4, 1.3], "D": [2.4, 2.0, 1.8, 1.6, 1.5], "E": [3.5, 3.2, 2.8, 2.4, 2.4] }; const y_arr = grid[siteClass] || grid["D"]; return interpolateLine(S1, [0.1, 0.2, 0.3, 0.4, 0.5], y_arr); }
function interpolateLine(x, x_arr, y_arr) { if (x <= x_arr[0]) return y_arr[0]; if (x >= x_arr[x_arr.length - 1]) return y_arr[y_arr.length - 1]; for (let i = 0; i < x_arr.length - 1; i++) { if (x >= x_arr[i] && x <= x_arr[i+1]) return y_arr[i] + (x - x_arr[i]) * (y_arr[i+1] - y_arr[i]) / (x_arr[i+1] - x_arr[i]); } return y_arr[0]; }
function sNum(val, decimals) { return (typeof val === 'number' && !isNaN(val)) ? val.toFixed(decimals) : (val ? val : "-"); }

// --- REPORT GENERATION ---
function createCalculationSheet(res) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.sheetCalc);
  if (!sheet) sheet = ss.insertSheet(CONFIG.sheetCalc); else sheet.clear();
  
  const p = res.params; const levels = res.levels.slice().reverse();
  sheet.setColumnWidths(1, 10, 100);
  
  let r = 1;
  sheet.getRange(r, 1).setValue("SEISMIC ANALYSIS REPORT").setFontSize(14).setFontWeight("bold"); r += 1;
  sheet.getRange(r, 1).setValue("Method: " + p.meth).setFontSize(10).setFontStyle("italic").setFontColor("#555"); r += 2;

  let paramsData = [];
  if (p.basinZone > 0) {
    paramsData = [ ["Parameter", "Value", "Description"], ["Zone", p.basinZone, "Bangkok Basin Zone"], ["Site Class", "E", "Soft Clay"], ["T (Period)", sNum(p.T, 3) + " s", "Fundamental Period"], ["Sa (0.5s)", sNum(p.Sa05, 3), "Spectral Acc. at 0.5s"], ["SDC", "Type " + p.dc, "Seismic Design Category"], ["System", p.systemType, "Selected System"], ["R", p.R, "Response Mod. Factor"], ["Ie", p.Ie, "Importance Factor"], ["Cs", sNum(p.Cs, 4), "Calculated from Sa(0.5s)"], ["V (Base Shear)", sNum(p.V, 2) + " T", "Design Base Shear"] ];
  } else {
    paramsData = [ ["Parameter", "Value", "Description"], ["Ss", sNum(p.Ss, 3), "Spectral Acc. (Short)"], ["S1", sNum(p.S1, 3), "Spectral Acc. (1.0s)"], ["Site Class", p.SiteClass || "-", "Soil Type"], ["Fa", sNum(p.Fa, 2), "Site Coeff. (Short)"], ["Fv", sNum(p.Fv, 2), "Site Coeff. (Long)"], ["SDS", sNum(p.SDS, 3), "Design Spectral Acc. (Short)"], ["SD1", sNum(p.SD1, 3), "Design Spectral Acc. (1.0s)"], ["T (Period)", sNum(p.T, 3) + " s", "Fundamental Period"], ["SDC", "Type " + p.dc, "Seismic Design Category"], ["System", p.systemType, "Selected System"], ["R", p.R, "Response Mod. Factor"], ["Ie", p.Ie, "Importance Factor"], ["Cs", sNum(p.Cs, 4), "Seismic Response Coeff."], ["V (Base Shear)", sNum(p.V, 2) + " T", "Design Base Shear"] ];
  }
  sheet.getRange(r, 1, paramsData.length, 3).setValues(paramsData).setBorder(true, true, true, true, true, true, "#999999", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(r, 1, 1, 3).setBackground(CONFIG.colors.tableHeader).setFontColor("white").setFontWeight("bold");
  sheet.getRange(r, 2, paramsData.length, 1).setHorizontalAlignment("left");
  r += paramsData.length + 1;
  
  if (p.isAllowed === false) { sheet.getRange(r, 1, 2, 7).merge().setValue("❌ WARNING: " + p.violationMsg).setBackground("#fce8e6").setFontColor("#c5221f").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle"); r += 3; } else { r += 1; }

  sheet.getRange(r, 1).setValue("1. VERTICAL DISTRIBUTION OF FORCES").setFontWeight("bold"); r++;
  let head1 = [["Story", "Height hx", "Weight wx", "wx * hx^k", "Cvx", "Force Fx", "Shear Vx"]];
  sheet.getRange(r, 1, 1, 7).setValues(head1).setBackground(CONFIG.colors.tableHeader).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center"); r++;
  let table1 = [];
  for (let i = 0; i < levels.length; i++) {
    let l = levels[i]; let cvx = (p.V > 0) ? (l.Fx / p.V) : 0;
    table1.push([ l.name, sNum(l.hx, 2), sNum(l.wx, 2), sNum(l.wx_hx_k, 1), sNum(cvx, 4), sNum(l.Fx, 3), l.Vx ? sNum(l.Vx, 3) : "-" ]);
  }
  sheet.getRange(r, 1, table1.length, 7).setValues(table1).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment("center");
  for (let i = 0; i < table1.length; i++) if (i % 2 !== 0) sheet.getRange(r + i, 1, 1, 7).setBackground(CONFIG.colors.tableRowEven);
  r += table1.length + 2;

  function writeStabilityTable(title, isX) {
    sheet.getRange(r, 1).setValue(title).setFontWeight("bold"); r++;
    let head2 = [["Story", "kx (T/m)", "Px (T)", "Δx (mm)", "Drift Ratio", "Drift Stat", "P-Delta θ", "P-Delta Stat", "Overturn Mx"]];
    sheet.getRange(r, 1, 1, 9).setValues(head2).setBackground("#00695c").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center"); r++;
    let table2 = [];
    for (let i = 0; i < levels.length; i++) {
      let l = levels[i];
      if (l.hx > 0) { table2.push([ l.name, isX ? sNum(l.kx_X, 0) : sNum(l.kx_Y, 0), sNum(l.Px, 1), isX ? sNum(l.dx_X, 2) : sNum(l.dx_Y, 2), isX ? sNum(l.drift_X, 5) : sNum(l.drift_Y, 5), isX ? (l.dStat_X || "-") : (l.dStat_Y || "-"), isX ? sNum(l.th_X, 4) : sNum(l.th_Y, 4), isX ? (l.pStat_X || "-") : (l.pStat_Y || "-"), sNum(l.Mx, 2) ]); }
    }
    if (table2.length > 0) {
      sheet.getRange(r, 1, table2.length, 9).setValues(table2).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment("center");
      for (let i = 0; i < table2.length; i++) {
        if (i % 2 !== 0) sheet.getRange(r + i, 1, 1, 9).setBackground(CONFIG.colors.tableRowEven);
        let dStat = table2[i][5]; let pStat = table2[i][7];
        if (dStat === "FAIL") sheet.getRange(r + i, 6).setBackground("#fce8e6").setFontColor("#c5221f").setFontWeight("bold"); else if (dStat === "OK") sheet.getRange(r + i, 6).setFontColor("#137333").setFontWeight("bold");
        if (pStat.includes("FAIL")) sheet.getRange(r + i, 8).setBackground("#fce8e6").setFontColor("#c5221f").setFontWeight("bold"); else if (pStat.includes("Amplify")) sheet.getRange(r + i, 8).setBackground("#fef7e0").setFontColor("#b06000").setFontWeight("bold"); else if (pStat === "OK") sheet.getRange(r + i, 8).setFontColor("#137333").setFontWeight("bold");
      }
      r += table2.length + 2;
    }
  }

  writeStabilityTable("2. STABILITY & DRIFT CHECKS (X-DIRECTION / E-W)", true);
  writeStabilityTable("3. STABILITY & DRIFT CHECKS (Y-DIRECTION / N-S)", false);

  sheet.getRange(r, 1).setValue("4. ORTHOGONAL SEISMIC FORCES (100/30 RULE)").setFontWeight("bold"); r++;
  let isMandatory = (p.dc === "C" || p.dc === "D" || p.dc === "ค (C)" || p.dc === "ง (D)");
  let sdcNote = isMandatory ? "⚠️ Mandatory: The 100-30 rule is REQUIRED for SDC " + p.dc + "." : "ℹ️ Optional: The 100-30 rule is generally NOT required for SDC " + p.dc + ".";
  let noteColor = isMandatory ? "#d93025" : "#1a73e8";
  sheet.getRange(r, 1, 1, 7).merge().setValue(sdcNote).setFontColor(noteColor).setFontStyle("italic").setHorizontalAlignment("left"); r++;
  let head3 = [["Story", "Fx (100%)", "Fy (100%)", "1.0Fx + 0.3Fy", "1.0Fx - 0.3Fy", "1.0Fy + 0.3Fx", "1.0Fy - 0.3Fx"]];
  sheet.getRange(r, 1, 1, 7).setValues(head3).setBackground("#455a64").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center"); r++;
  let table3 = [];
  for (let i = 0; i < levels.length; i++) { if (levels[i].hx > 0) { let f = levels[i].Fx || 0; table3.push([ levels[i].name, sNum(f, 3), sNum(f, 3), sNum(1.3 * f, 3), sNum(0.7 * f, 3), sNum(1.3 * f, 3), sNum(0.7 * f, 3) ]); } }
  if (table3.length > 0) { sheet.getRange(r, 1, table3.length, 7).setValues(table3).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment("center"); for (let i = 0; i < table3.length; i++) if (i % 2 !== 0) sheet.getRange(r + i, 1, 1, 7).setBackground(CONFIG.colors.tableRowEven); }
}

// --- DRAWING LOGIC (DUAL-AXIS PLAN) ---
function generateBlueprintFromData(rawSpanX, rawHeight, rawLoads, rawSpanY, rawPointLoads, calcResult) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let planSheet = ss.getSheetByName(CONFIG.sheetPlan);
  if (!planSheet) planSheet = ss.insertSheet(CONFIG.sheetPlan);
  
  // ล้างการตั้งค่าเก่าทั้งหมดรวมถึงการ Break Apart Merge Cells เก่าทิ้ง ป้องกันบั๊กการทับซ้อน
  planSheet.clear();
  try { planSheet.getRange(1, 1, planSheet.getMaxRows(), planSheet.getMaxColumns()).breakApart(); } catch(e) {}

  const spansX_meters = rawSpanX.split(',').map(Number);
  const spansY_meters = rawSpanY.split(',').map(Number);
  const heights_meters = rawHeight.split(',').map(Number).reverse();
  const loads_val = rawLoads.split(',').map(Number);
  const ptLoads_val = rawPointLoads.split(',').map(Number);

  let minSpan = Math.min(...spansX_meters, ...spansY_meters);
  let res = CONFIG.defaultResolution;
  if (minSpan < 1.5) res = 0.125; else if (minSpan < 3.0) res = 0.25;

  const spansX_cells = spansX_meters.map(n => Math.round((n || 0) / res));
  const spansY_cells = spansY_meters.map(n => Math.round((n || 0) / res));
  const heights_cells = heights_meters.map(n => Math.round((n || 0) / res));

  let totalWidthX_cells = 0; for (let i = 0; i < spansX_cells.length; i++) totalWidthX_cells += spansX_cells[i];
  let totalWidthY_cells = 0; for (let i = 0; i < spansY_cells.length; i++) totalWidthY_cells += spansY_cells[i];
  let totalHeight_cells = 0; for (let i = 0; i < heights_cells.length; i++) totalHeight_cells += heights_cells[i];

  let maxF = 0; for (let i = 0; i < ptLoads_val.length; i++) if (ptLoads_val[i] > maxF) maxF = ptLoads_val[i];
  let maxV = (calcResult && calcResult.params && calcResult.params.V) ? calcResult.params.V : 0;

  let largestWidth = Math.max(totalWidthX_cells, totalWidthY_cells);
  let dScale = (104 - largestWidth - 20) / ((maxF + maxV) || 1);
  if (dScale > 3.0) dScale = 3.0; if (dScale < 0.2) dScale = 0.2;

  // [FORMAT UPDATE] กำหนดระยะขอบซ้ายให้พอดีกับลูกศร เพื่อดันแปลนไปชิดซ้ายสุด
  const requiredLeftSpace = Math.ceil(maxF * dScale) + 20; 
  const requiredRightSpace = Math.ceil(maxV * dScale) + 20;

  const canvasWidth = Math.max(largestWidth + requiredLeftSpace + requiredRightSpace + 40, 80);
  const totalRowsNeeded = (totalHeight_cells * 2) + totalWidthX_cells + totalWidthY_cells + 100;

  if (canvasWidth > planSheet.getMaxColumns()) planSheet.insertColumnsAfter(planSheet.getMaxColumns(), canvasWidth - planSheet.getMaxColumns());
  if (totalRowsNeeded > planSheet.getMaxRows()) planSheet.insertRowsAfter(planSheet.getMaxRows(), totalRowsNeeded - planSheet.getMaxRows());

  planSheet.setColumnWidths(1, canvasWidth, CONFIG.cellSizePx);
  planSheet.setRowHeights(1, totalRowsNeeded, CONFIG.cellSizePx);
  planSheet.getRange(1, 1, totalRowsNeeded, canvasWidth).setBorder(true, true, true, true, true, true, CONFIG.colors.graphLine, SpreadsheetApp.BorderStyle.DOTTED);

  function drawSet(startRow, spansHoriz_m, spansHoriz_c, spansDepth_m, spansDepth_c, axisName, loadPrefix) {
    let drawingWidth = 0; for (let i = 0; i < spansHoriz_c.length; i++) drawingWidth += spansHoriz_c[i];
    
    // เริ่มวาดที่คอลัมน์ด้านซ้าย
    let startCol = requiredLeftSpace;
    if (startCol < 15) startCol = 15; 
    
    let currentRow = startRow;
    
    planSheet.getRange(currentRow - 6, startCol).setValue("ELEVATION VIEW (" + axisName + " Axis)").setFontSize(12).setFontWeight("bold").setFontColor("#1565c0");
    
    for (let idx = 0; idx < heights_cells.length; idx++) {
      let currentX = startCol;
      let h_cells = heights_cells[idx];
      let h_meters = heights_meters[idx] || 0;
      
      let pLoadIndex = (ptLoads_val.length - 1) - idx;
      let dLoadIndex = (loads_val.length - 1) - idx;
      let pLoad = ptLoads_val[pLoadIndex] || 0;
      let dLoad = loads_val[dLoadIndex] || 0;
      
      // [FIX] ย่อความสูงของ Merge Cell ตัวเลขระยะความสูง เพื่อไม่ให้ล้ำไปชนกับแกนของลูกศร (แถวแนวนอนของคาน)
      let dimHeight = h_cells - 2; 
      if (dimHeight < 1) dimHeight = 1;
      planSheet.getRange(currentRow + 1, startCol - 4, dimHeight, 3).merge()
        .setValue(h_meters.toFixed(2) + "m")
        .setFontColor(CONFIG.colors.dimText).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(8);
      
      if (pLoad > 0) drawLateralLoad(planSheet, currentRow, startCol, pLoad, dScale, loadPrefix);
      if (calcResult && calcResult.levels) {
        let shearIndex = (calcResult.levels.length - 1) - idx;
        let shearVal = calcResult.levels[shearIndex] ? calcResult.levels[shearIndex].Vx : 0;
        if (shearVal > 0) drawShearForceRight(planSheet, currentRow, startCol + drawingWidth, h_cells, shearVal, dScale);
      }
      
      for (let i = 0; i < spansHoriz_c.length; i++) {
        let w_cells = spansHoriz_c[i];
        let room = planSheet.getRange(currentRow, currentX, h_cells, w_cells);
        room.setBackground(CONFIG.colors.fillSide).setBorder(true, true, true, true, null, null, CONFIG.colors.beam, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        
        if (dLoad > 0) {
          drawLoadArrows(planSheet, currentRow, currentX, w_cells);
          if (i === 0) drawLoadLabel(planSheet, currentRow, currentX, w_cells, dLoad);
        }
        currentX += w_cells;
      }
      createLabelBox(planSheet, currentRow + 1, currentX - 3, "FL " + (heights_cells.length - idx), CONFIG.colors.gridLabel);
      currentRow += h_cells;
    }

    let basePointLoad = ptLoads_val[0] || 0;
    if (basePointLoad > 0) drawLateralLoad(planSheet, currentRow, startCol, basePointLoad, dScale, loadPrefix);
    
    let baseDistLoad = loads_val[0] || 0;
    if (baseDistLoad > 0) {
      let currentX = startCol;
      for (let i = 0; i < spansHoriz_c.length; i++) {
        drawLoadArrows(planSheet, currentRow, currentX, spansHoriz_c[i]);
        if (i === 0) drawLoadLabel(planSheet, currentRow, currentX, spansHoriz_c[i], baseDistLoad);
        currentX += spansHoriz_c[i];
      }
    }

    let columnX = startCol;
    let colPositions = [startCol];
    for (let i = 0; i < spansHoriz_c.length; i++) { columnX += spansHoriz_c[i]; colPositions.push(columnX); }
    for (let i = 0; i < colPositions.length; i++) {
      drawColumnStump(planSheet, currentRow, colPositions[i], CONFIG.stumpHeight);
      drawFixedSupport(planSheet, currentRow + CONFIG.stumpHeight, colPositions[i]);
    }

    let gridX = startCol; let gridNum = 1; let labelRow = currentRow + CONFIG.stumpHeight + 4;
    createLabelBox(planSheet, labelRow, gridX - 1, gridNum, CONFIG.colors.gridLabel); gridNum++;
    
    // [FIX] ลบขอบซ้ายขวาของ Merge Cell ระยะความกว้าง เพื่อไม่ให้เกยทับกับตัวเลข Grid สีแดง (1, 2, 3...)
    for (let i = 0; i < spansHoriz_c.length; i++) {
      let spaceWidth = spansHoriz_c[i] - 2; 
      if (spaceWidth > 0) {
        planSheet.getRange(labelRow, gridX + 1, 2, spaceWidth).merge()
          .setValue((spansHoriz_m[i] || 0).toFixed(2) + "m")
          .setFontColor(CONFIG.colors.dimText).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(8);
      }
      gridX += spansHoriz_c[i];
      createLabelBox(planSheet, labelRow, gridX - 1, gridNum, CONFIG.colors.gridLabel); gridNum++;
    }
    
    // Total Span Length
    let totalSpanM = 0; for (let i = 0; i < spansHoriz_m.length; i++) totalSpanM += (spansHoriz_m[i] || 0);
    let totLabelRow = labelRow + 3;
    planSheet.getRange(totLabelRow, startCol, 1, drawingWidth).setBorder(false, false, true, false, false, false, CONFIG.colors.dimText, SpreadsheetApp.BorderStyle.SOLID);
    planSheet.getRange(totLabelRow + 1, startCol, 2, drawingWidth).merge()
      .setValue("Total = " + totalSpanM.toFixed(2) + "m")
      .setFontColor(CONFIG.colors.dimText).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontWeight("bold").setFontSize(9);

    // --- 2. PLAN VIEW ---
    currentRow += 16;
    planSheet.getRange(currentRow - 4, startCol).setValue("PLAN VIEW (" + axisName + " Axis)").setFontSize(12).setFontWeight("bold").setFontColor("#1565c0");
    
    gridX = startCol; gridNum = 1;
    createLabelBox(planSheet, currentRow - 3, gridX - 1, gridNum, CONFIG.colors.gridLabel); gridNum++;
    
    for (let i = 0; i < spansHoriz_c.length; i++) {
      let spaceWidth = spansHoriz_c[i] - 2;
      if (spaceWidth > 0) {
        planSheet.getRange(currentRow - 3, gridX + 1, 2, spaceWidth).merge()
          .setValue((spansHoriz_m[i] || 0).toFixed(2) + "m")
          .setFontColor(CONFIG.colors.dimText).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(8);
      }
      gridX += spansHoriz_c[i];
      createLabelBox(planSheet, currentRow - 3, gridX - 1, gridNum, CONFIG.colors.gridLabel); gridNum++;
    }
    
    for (let idx = 0; idx < spansDepth_c.length; idx++) {
      let currentX = startCol; let hc = spansDepth_c[idx];
      let charCode = String.fromCharCode(65 + idx);
      
      createLabelBox(planSheet, currentRow - 1, startCol - 3, charCode, CONFIG.colors.gridLabel);
      
      // [FIX] ย่อความสูง Dimension บนแบบแปลน (Plan view) เพื่อไม่ให้ชนกับเส้นกริดแนวนอน
      let dimHeight = hc - 2;
      if (dimHeight < 1) dimHeight = 1;
      planSheet.getRange(currentRow + 1, startCol - 4, dimHeight, 3).merge()
        .setValue((spansDepth_m[idx] || 0).toFixed(2) + "m")
        .setFontColor(CONFIG.colors.dimText).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(8);
      
      for (let i = 0; i < spansHoriz_c.length; i++) {
        let room = planSheet.getRange(currentRow, currentX, hc, spansHoriz_c[i]);
        room.setBackground(CONFIG.colors.fillTop).setBorder(true, true, true, true, null, null, "#90a4ae", SpreadsheetApp.BorderStyle.SOLID);
        currentX += spansHoriz_c[i];
      }
      
      if (idx === spansDepth_c.length - 1) {
        createLabelBox(planSheet, currentRow + hc - 1, startCol - 3, String.fromCharCode(65 + idx + 1), CONFIG.colors.gridLabel);
      }
      currentRow += hc;
    }
    return currentRow + 10;
  }

  let nextRow = drawSet(8, spansX_meters, spansX_cells, spansY_meters, spansY_cells, "E-W / X", "Fx");
  drawSet(nextRow + 5, spansY_meters, spansY_cells, spansX_meters, spansX_cells, "N-S / Y", "Fy");
  
  // SECTION DETAILS (Text Only)
  if (calcResult && calcResult.params && calcResult.params.stiffnessConfig) {
    const sc = calcResult.params.stiffnessConfig;
    
    let secCol = requiredLeftSpace + largestWidth + requiredRightSpace + 5;
    if (secCol > canvasWidth - 10) secCol = canvasWidth - 10; 
    
    planSheet.getRange(4, secCol, 1, 10).merge().setValue("SECTION DETAILS").setFontWeight("bold").setFontSize(10).setFontColor("#1565c0").setHorizontalAlignment("left");
    
    planSheet.getRange(6, secCol, 1, 10).merge().setValue("• Column: " + sNum(sc.col_b, 2) + " x " + sNum(sc.col_h, 2) + " m.").setFontSize(9).setFontWeight("bold").setHorizontalAlignment("left");
    planSheet.getRange(8, secCol, 1, 10).merge().setValue("• Beam X: " + sNum(sc.bmX_b, 2) + " x " + sNum(sc.bmX_h, 2) + " m.").setFontSize(9).setFontWeight("bold").setHorizontalAlignment("left");
    planSheet.getRange(10, secCol, 1, 10).merge().setValue("• Beam Y: " + sNum(sc.bmY_b, 2) + " x " + sNum(sc.bmY_h, 2) + " m.").setFontSize(9).setFontWeight("bold").setHorizontalAlignment("left");
  }

  planSheet.setHiddenGridlines(true);
}

// [FORMAT UPDATE] ให้ลูกศรกระเถิบเข้าไปชิดผนังมากที่สุด และขยายกล่องข้อความให้แสดงผลครบ
function drawShearForceRight(sheet, row, col, height, val, scale) {
  let arrowLength = Math.max(3, Math.ceil(val * scale));
  let line = ""; for (let i = 0; i < Math.max(1, arrowLength); i++) line += "─";
  let text = "←" + line + " V = " + (val || 0).toFixed(2) + "T";
  
  sheet.getRange(row + Math.floor(height / 2) - 1, col, 2, arrowLength + 15).merge()
    .setValue(text).setHorizontalAlignment("left").setVerticalAlignment("middle")
    .setFontColor(CONFIG.colors.shearArrow).setFontWeight("bold").setFontSize(9);
}

function drawLateralLoad(sheet, beamRow, startCol, val, scale, prefix) {
  let arrowLength = Math.max(3, Math.ceil(val * scale));
  let rightGap = 1; // [FIX] เปลี่ยนจาก 8 เป็น 1 เพื่อให้หัวลูกศรชนกับผนังอาคารพอดี
  let textSpace = 14; // ให้พื้นที่ตัวหนังสือเยอะขึ้นเพื่อป้องกันการล้น
  
  let targetCol = startCol - arrowLength - rightGap - textSpace;
  if (targetCol < 1) targetCol = 1;
  
  let line = ""; for (let i = 0; i < Math.max(1, arrowLength); i++) line += "─";
  let valStr = (typeof val === 'number') ? val.toFixed(3) : val;
  let text = prefix + " = " + valStr + "T " + line + "→";
  
  sheet.getRange(beamRow - 1, targetCol, 2, arrowLength + textSpace).merge()
    .setValue(text).setHorizontalAlignment("right").setVerticalAlignment("middle")
    .setFontColor(CONFIG.colors.loadArrow).setFontWeight("bold").setFontSize(9);
}

function drawLoadArrows(sheet, beamRow, startCol, width) {
  if (beamRow - 1 > 0) {
    let arrows = ""; let count = Math.max(1, Math.floor(width - 1));
    for (let i = 0; i < count; i++) arrows += "↓ ";
    sheet.getRange(beamRow - 1, startCol, 1, width).merge().setValue(arrows).setHorizontalAlignment("center").setVerticalAlignment("bottom").setFontColor(CONFIG.colors.loadArrow).setFontSize(8).setFontWeight("bold");
  }
}

function drawLoadLabel(sheet, beamRow, startCol, width, val) {
  if (beamRow - 3 > 0) {
    sheet.getRange(beamRow - 3, startCol, 2, width).merge().setValue(val + " T/m²").setHorizontalAlignment("center").setVerticalAlignment("middle").setFontColor(CONFIG.colors.loadText).setFontSize(9).setFontWeight("bold");
  }
}

function createLabelBox(sheet, row, col, text, color, align, isBold, fontSize) {
  if (row < 1 || col < 1) return;
  if (align === undefined) align = "center"; if (isBold === undefined) isBold = true; if (fontSize === undefined) fontSize = 8;
  let range = sheet.getRange(row, col, 2, 2);
  range.merge().setValue(text).setFontColor(color).setBackground(CONFIG.colors.labelBg).setBorder(true, true, true, true, null, null, "#dddddd", SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment(align).setVerticalAlignment("middle").setFontSize(fontSize);
  if (isBold) range.setFontWeight("bold");
}

function drawFixedSupport(sheet, row, centerX) {
  let startX = centerX - 2;
  if (startX > 0) {
    sheet.getRange(row, startX, 2, 4).merge().setBackground(CONFIG.colors.support).setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.getRange(row + 2, startX - 1, 1, 6).merge().setValue("/ / / / / /").setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(8).setFontColor("#757575").setFontStyle("italic");
  }
}

function drawColumnStump(sheet, row, x, height) {
  sheet.getRange(row, x, height, 1).setBorder(null, true, null, null, null, null, CONFIG.colors.beam, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}