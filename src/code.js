/**
 * Building Structure Generator & Seismic Analyzer - V0.5.34
 * + UI FIX: Reduced `requiredLeftSpace` from 45 to 36 to eliminate the large gap between the dimension line and the building drawing.
 * + FEATURE: Integrated "Project Info" dynamically with Cover Page and Frozen Headers.
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
    beam: "#37474f", fillSide: "#f3e5f5", fillTop: "#e1f5fe",
    gridLabel: "#b71c1c", dimText: "#0d47a1", graphLine: "#eceff1",
    labelBg: "#ffffff", support: "#424242",
    loadArrow: "#b71c1c", loadText: "#b71c1c",
    shearArrow: "#00695c", shearText: "#004d40",
    tableHeader: "#1565c0", tableRowEven: "#f5f5f5",
    warningText: "#d93025", passText: "#137333", failText: "#c5221f"
  }
};

const BANGKOK_ZONE_DATA = {
  1: {Sa05:0.360}, 2: {Sa05:0.352}, 3: {Sa05:0.262}, 4: {Sa05:0.287}, 5: {Sa05:0.191}, 
  6: {Sa05:0.272}, 7: {Sa05:0.246}, 8: {Sa05:0.162}, 9: {Sa05:0.214}, 10: {Sa05:0.179}
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
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('RC Seismic Analysis').setWidth(350);
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

function calculateELF(spanXStr, heightStr, loadsStr, spanYStr, params) {
  const spanX = (spanXStr || "").toString().split(',').map(Number);
  const spanY = (spanYStr || "").toString().split(',').map(Number);
  const heights = (heightStr || "").toString().split(',').map(Number).filter(n => n > 0); 
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
  let designCategory = ""; let calculationMethod = ""; let Cs = 0;

  if (params.basinZone > 0) {
    calculationMethod = "Bangkok Basin (Zone " + params.basinZone + ") - Using Sa(0.5s)";
    const zoneData = BANGKOK_ZONE_DATA[params.basinZone];
    Sa05 = zoneData ? zoneData.Sa05 : 0;
    if (T <= 0.5) { designCategory = classifySDC_Table223(Sa05, RiskCat); calculationMethod += " | T=" + T.toFixed(2) + "s (<=0.5s) -> Use Table 2.2.3"; } 
    else { designCategory = classifySDC_Table224(Sa05, RiskCat); calculationMethod += " | T=" + T.toFixed(2) + "s (>0.5s) -> Use Table 2.2.4"; }
    SDS = Sa05; SD1 = 0; Cs = Sa05 / (R / Ie); 
  } else {
    calculationMethod = "Standard ELF (DPT 1301/1302)";
    Fa = getFa(SiteClass, Ss); Fv = getFv(SiteClass, S1);
    SDS = (2/3) * Fa * Ss; SD1 = (2/3) * Fv * S1;
    const catShort = classifySDC_Table223(SDS, RiskCat); const catLong = classifySDC_Table224(SD1, RiskCat);
    designCategory = (catShort > catLong) ? catShort : catLong;
    Cs = SDS / (R / Ie);
    const Cs_max = SD1 / (T * (R / Ie));
    if (Cs > Cs_max) Cs = Cs_max;
    if (Ss === 0 && S1 === 0) Cs = 0; else if (Cs < 0.01) Cs = 0.01;
  }

  let totalSpanLength = 0; for (let i = 0; i < spanX.length; i++) totalSpanLength += (spanX[i] || 0);
  let totalDepth = 0; for (let i = 0; i < spanY.length; i++) totalDepth += (spanY[i] || 0);

  let floorArea = totalSpanLength * totalDepth;
  let numFloors = heights.length + 1;
  let totalArea = floorArea * numFloors;

  let levels = []; 
  let currentHeight = 0; 
  levels.push({ name: "Base/FL1", hx: 0, wx: (loads[0] || 0) * floorArea, wx_hx_k: 0, Fx: 0, Vx: 0 });

  for (let i = 0; i < heights.length; i++) {
    currentHeight += (heights[i] || 0);
    const loadIndex = i + 1;
    const isRoof = (i === heights.length - 1);
    levels.push({ name: isRoof ? "Roof" : "FL " + (i + 2), hx: currentHeight, wx: (loads[loadIndex] || 0) * floorArea, wx_hx_k: 0 });
  }

  let W_effective = 0;
  for (let i = 0; i < levels.length; i++) if (levels[i].hx > 0) W_effective += levels[i].wx;
  
  let W_total = 0;
  for (let i = 0; i < levels.length; i++) {
    if (levels[i].hx > 0) W_total += levels[i].wx;
  }

  const V = Cs * W_effective;
  let k = 1; if (T <= 0.5) k = 1; else if (T >= 2.5) k = 2; else k = 1 + ((T - 0.5) / 2);

  let sum_w_h_k = 0;
  for (let i = 0; i < levels.length; i++) {
    if (levels[i].hx > 0) { levels[i].wx_hx_k = levels[i].wx * Math.pow(levels[i].hx, k); sum_w_h_k += levels[i].wx_hx_k; }
  }

  let pointLoadsArr = [];
  for (let i = 0; i < levels.length; i++) {
    if (levels[i].hx > 0 && sum_w_h_k > 0) levels[i].Fx = (levels[i].wx_hx_k / sum_w_h_k) * V;
    else levels[i].Fx = 0;
    pointLoadsArr.push((levels[i].Fx || 0).toFixed(3));
  }

  let cumShear = 0;
  for (let i = levels.length - 1; i >= 0; i--) { cumShear += levels[i].Fx; levels[i].Vx = cumShear; }

  return { 
    pointLoadsStr: pointLoadsArr.join(","), levels: levels, 
    params: { 
      province: params.province, amphoe: params.amphoe, 
      Sa05: Sa05, T: T, Cs: Cs, V: V, Weff: W_effective, W_total: W_total, 
      Ss: Ss, S1: S1, SDS: SDS, SD1: SD1, Fa: Fa, Fv: Fv, SiteClass: SiteClass, RiskCat: RiskCat, 
      R: R, Cd: params.Cd, Ie: Ie, dc: designCategory, meth: calculationMethod, 
      basinZone: params.basinZone, systemType: params.systemType, isAllowed: params.isAllowed, violationMsg: params.violationMsg,
      stiffnessConfig: params.stiffnessConfig, gravConfig: params.gravConfig, projectInfo: params.projectInfo,
      totalSpanX: totalSpanLength, totalSpanY: totalDepth, totalHeight: totalHeight,
      floorArea: floorArea, numFloors: numFloors, totalArea: totalArea, 
      spanXStr: spanXStr, spanYStr: spanYStr 
    } 
  };
}

function calculateStability(calcResult, params, spanXStr, spanYStr, heightsStr) {
  const levels = calcResult.levels; const sParams = params.stiffnessConfig;
  if (!sParams) return calcResult; 

  const spanX = (spanXStr || "").toString().split(',').map(Number);
  const spanY = (spanYStr || "").toString().split(',').map(Number);
  const heights = (heightsStr || "").toString().split(',').map(Number).filter(n => n > 0);
  
  const nx = spanX.length; const ny = spanY.length; const numCols = (nx + 1) * (ny + 1);
  let sum_1_Lx = 0; for (let i = 0; i < spanX.length; i++) if(spanX[i]) sum_1_Lx += (1 / spanX[i]);
  let sum_1_Ly = 0; for (let i = 0; i < spanY.length; i++) if(spanY[i]) sum_1_Ly += (1 / spanY[i]);
  
  const frameFactor_Ib_Lx = (ny + 1) * sum_1_Lx; const frameFactor_Ib_Ly = (nx + 1) * sum_1_Ly; 
  const Ec = 15100 * Math.sqrt(sParams.fc || 280) * 10; 
  
  let Cd_val = parseFloat(params.Cd);
  if (isNaN(Cd_val) || Cd_val <= 0) Cd_val = 2.5;
  const Ie = parseFloat(params.Ie) || 1.0;
  
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
    let delta_x_X = delta_xe_X * (Cd_val / Ie); let driftRatio_X = delta_x_X / hs;
    
    let Ig_cx = (col_b * Math.pow(col_h, 3)) / 12; let Ig_by = (byb * Math.pow(byh, 3)) / 12;
    let stiffness_col_Y = numCols * ((0.70 * Ig_cx) / hs); let stiffness_bm_Y = frameFactor_Ib_Ly * (0.35 * Ig_by);
    let kx_Y = (12 * Ec) / (hs * hs * ((1 / stiffness_col_Y) + (1 / stiffness_bm_Y)));
    let delta_xe_Y = (kx_Y > 0) ? (Vx / kx_Y) : 0;
    let delta_x_Y = delta_xe_Y * (Cd_val / Ie); let driftRatio_Y = delta_x_Y / hs;
    
    let theta_max = 0.5 / Cd_val; if (theta_max > 0.25) theta_max = 0.25;
    let theta_X = 0; let pStatus_X = "OK";
    if (Vx > 0 && hs > 0 && Cd_val > 0) { 
      theta_X = (Px * delta_x_X) / (Vx * hs * Cd_val); 
      if (theta_X > theta_max) pStatus_X = "FAIL"; 
      else if (theta_X > 0.1) pStatus_X = "Amplify"; 
    }

    let theta_Y = 0; let pStatus_Y = "OK";
    if (Vx > 0 && hs > 0 && Cd_val > 0) { 
      theta_Y = (Px * delta_x_Y) / (Vx * hs * Cd_val); 
      if (theta_Y > theta_max) pStatus_Y = "FAIL"; 
      else if (theta_Y > 0.1) pStatus_Y = "Amplify"; 
    }

    let Mx = 0; let h_base = levels[i-1].hx; for (let j = i; j < levels.length; j++) Mx += levels[j].Fx * (levels[j].hx - h_base);
    
    levels[i].Px = Px; levels[i].Mx = Mx; levels[i].theta_max = theta_max;
    levels[i].kx_X = kx_X; levels[i].dx_X = delta_x_X * 1000; levels[i].drift_X = driftRatio_X; levels[i].dStat_X = (driftRatio_X <= 0.01) ? "OK" : "FAIL"; levels[i].th_X = theta_X; levels[i].pStat_X = pStatus_X;
    levels[i].kx_Y = kx_Y; levels[i].dx_Y = delta_x_Y * 1000; levels[i].drift_Y = driftRatio_Y; levels[i].dStat_Y = (driftRatio_Y <= 0.01) ? "OK" : "FAIL"; levels[i].th_Y = theta_Y; levels[i].pStat_Y = pStatus_Y;
  }
  return calcResult;
}

function classifySDC_Table223(val, riskCat) { let rLvl = 0; if (riskCat === 'III') rLvl = 1; if (riskCat === 'IV') rLvl = 2; const cats = ["A", "B", "C", "D"]; if (val < 0.167) return cats[0]; if (val < 0.33) return cats[(rLvl === 2) ? 2 : 1]; if (val < 0.50) return cats[(rLvl === 2) ? 3 : 2]; return cats[3]; }
function classifySDC_Table224(val, riskCat) { let rLvl = 0; if (riskCat === 'III') rLvl = 1; if (riskCat === 'IV') rLvl = 2; const cats = ["A", "B", "C", "D"]; if (val < 0.067) return cats[0]; if (val < 0.133) return cats[(rLvl === 2) ? 2 : 1]; if (val < 0.20) return cats[(rLvl === 2) ? 3 : 2]; return cats[3]; }
function getFa(siteClass, Ss) { const grid = { "A": [0.8, 0.8, 0.8, 0.8, 0.8], "B": [1.0, 1.0, 1.0, 1.0, 1.0], "C": [1.2, 1.2, 1.1, 1.0, 1.0], "D": [1.6, 1.4, 1.2, 1.1, 1.0], "E": [2.5, 1.7, 1.2, 0.9, 0.9] }; const y_arr = grid[siteClass] || grid["D"]; return interpolateLine(Ss, [0.25, 0.5, 0.75, 1.0, 1.25], y_arr); }
function getFv(siteClass, S1) { const grid = { "A": [0.8, 0.8, 0.8, 0.8, 0.8], "B": [1.0, 1.0, 1.0, 1.0, 1.0], "C": [1.7, 1.6, 1.5, 1.4, 1.3], "D": [2.4, 2.0, 1.8, 1.6, 1.5], "E": [3.5, 3.2, 2.8, 2.4, 2.4] }; const y_arr = grid[siteClass] || grid["D"]; return interpolateLine(S1, [0.1, 0.2, 0.3, 0.4, 0.5], y_arr); }
function interpolateLine(x, x_arr, y_arr) { if (x <= x_arr[0]) return y_arr[0]; if (x >= x_arr[x_arr.length - 1]) return y_arr[y_arr.length - 1]; for (let i = 0; i < x_arr.length - 1; i++) { if (x >= x_arr[i] && x <= x_arr[i+1]) return y_arr[i] + (x - x_arr[i]) * (y_arr[i+1] - y_arr[i]) / (x_arr[i+1] - x_arr[i]); } return y_arr[0]; }
function sNum(val, decimals) { return (typeof val === 'number' && !isNaN(val)) ? val.toFixed(decimals) : (val ? val : "-"); }
function fmtDate(dStr) { if(!dStr) return "-"; const p = dStr.split('-'); if(p.length===3) return p[2]+"/"+p[1]+"/"+p[0]; return dStr; }

function setStatus(range, statusText) {
  range.setValue(statusText); 
  if (statusText === "PASS" || statusText === "OK") {
    range.setFontColor(CONFIG.colors.passText).setFontWeight("bold").setBackground(null);
  } else if (statusText === "FAIL" || statusText.includes("FAIL")) {
    range.setFontColor(CONFIG.colors.failText).setFontWeight("bold").setBackground("#fce8e6");
  } else {
    range.setFontColor("#b06000").setFontWeight("bold").setBackground(null);
  }
}

function preparePdfExport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planSheet = ss.getSheetByName(CONFIG.sheetPlan);
  const reportSheet = ss.getSheetByName(CONFIG.sheetCalc);
  
  if (!planSheet || !reportSheet) throw new Error("ไม่พบหน้า Plan หรือ Calculation_Report กรุณากด Draw & Calculate ก่อน");
  
  const props = PropertiesService.getDocumentProperties();
  
  const planGapStart = parseInt(props.getProperty('planGapStart'));
  const planGapEnd = parseInt(props.getProperty('planGapEnd'));
  if (planGapStart && planGapEnd && planGapStart <= planGapEnd) {
     try { planSheet.showRows(planGapStart, planGapEnd - planGapStart + 1); } catch(e){}
  }

  const reportGapStart = parseInt(props.getProperty('reportGapStart'));
  const reportGapEnd = parseInt(props.getProperty('reportGapEnd'));
  if (reportGapStart && reportGapEnd && reportGapStart <= reportGapEnd) {
     try { reportSheet.showRows(reportGapStart, reportGapEnd - reportGapStart + 1); } catch(e){}
  }
  
  const dataSheet = ss.getSheetByName(CONFIG.sheetData);
  if (dataSheet) dataSheet.hideSheet();
  
  const maxCols = planSheet.getMaxColumns();
  if (maxCols >= 131) { planSheet.hideColumns(131, maxCols - 130); }
  
  SpreadsheetApp.flush(); 
  
  const ssId = ss.getId();
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
              "?exportFormat=pdf&format=pdf" +
              "&size=A4&portrait=true&fitw=true&gridlines=false" +
              "&ir=false&ic=false";
  return url;
}

function restorePdfState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planSheet = ss.getSheetByName(CONFIG.sheetPlan);
  const reportSheet = ss.getSheetByName(CONFIG.sheetCalc);
  const props = PropertiesService.getDocumentProperties();
  
  if (planSheet) {
    const planGapStart = parseInt(props.getProperty('planGapStart'));
    const planGapEnd = parseInt(props.getProperty('planGapEnd'));
    if (planGapStart && planGapEnd && planGapStart <= planGapEnd) {
       try { planSheet.hideRows(planGapStart, planGapEnd - planGapStart + 1); } catch(e){}
    }
  }

  if (reportSheet) {
    const reportGapStart = parseInt(props.getProperty('reportGapStart'));
    const reportGapEnd = parseInt(props.getProperty('reportGapEnd'));
    if (reportGapStart && reportGapEnd && reportGapStart <= reportGapEnd) {
       try { reportSheet.hideRows(reportGapStart, reportGapEnd - reportGapStart + 1); } catch(e){}
    }
  }
  SpreadsheetApp.flush();
}

function createCalculationSheet(res) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.sheetCalc);
  if (!sheet) sheet = ss.insertSheet(CONFIG.sheetCalc); 
  
  try { sheet.showRows(1, sheet.getMaxRows()); } catch(e) {}
  try { sheet.getDataRange().breakApart(); } catch(e) {} 
  sheet.clear();
  
  const p = res.params; 
  const sc = p.stiffnessConfig;
  const gc = p.gravConfig; 
  const pi = p.projectInfo || {}; // ข้อมูลโครงการ
  const levels = res.levels.slice().reverse();
  
  sheet.setColumnWidths(1, 1, 160); 
  sheet.setColumnWidths(2, 4, 70);  
  sheet.setColumnWidths(6, 1, 110); 
  sheet.setColumnWidths(7, 1, 70);  
  sheet.setColumnWidths(8, 1, 80);  
  sheet.setColumnWidths(9, 1, 70);   
  sheet.setColumnWidths(10, 1, 70);  

  sheet.setFrozenRows(4);
  sheet.getRange(1, 6, 1, 5).merge().setValue("Project: " + (pi.projName || "-")).setHorizontalAlignment("right").setFontWeight("bold").setFontColor("#555");
  sheet.getRange(2, 6, 1, 5).merge().setValue("Date: " + fmtDate(pi.projDate)).setHorizontalAlignment("right").setFontColor("#555");
  sheet.getRange(3, 6, 1, 5).merge().setValue("Designed By: " + (pi.desName || "-")).setHorizontalAlignment("right").setFontColor("#555");
  
  sheet.getRange(15, 1, 1, 10).merge().setValue("รายการคำนวณโครงสร้าง").setHorizontalAlignment("center").setFontSize(24).setFontWeight("bold");
  sheet.getRange(17, 1, 1, 10).merge().setValue("SEISMIC ANALYSIS REPORT").setHorizontalAlignment("center").setFontSize(18).setFontColor("#555555");
  
  sheet.getRange(22, 2, 1, 2).merge().setValue("ชื่อโครงการ:").setHorizontalAlignment("right").setFontWeight("bold");
  sheet.getRange(22, 4, 1, 6).merge().setValue(pi.projName || "-").setHorizontalAlignment("left");
  
  sheet.getRange(24, 2, 1, 2).merge().setValue("เจ้าของโครงการ:").setHorizontalAlignment("right").setFontWeight("bold");
  sheet.getRange(24, 4, 1, 6).merge().setValue(pi.owner || "-").setHorizontalAlignment("left");
  
  sheet.getRange(26, 2, 1, 2).merge().setValue("สถานที่ก่อสร้าง:").setHorizontalAlignment("right").setFontWeight("bold");
  sheet.getRange(26, 4, 1, 6).merge().setValue(pi.location || "-").setHorizontalAlignment("left");
  
  sheet.getRange(28, 2, 1, 2).merge().setValue("ประเภทอาคาร:").setHorizontalAlignment("right").setFontWeight("bold");
  sheet.getRange(28, 4, 1, 6).merge().setValue(pi.bldgType || "-").setHorizontalAlignment("left");
  
  sheet.getRange(34, 2, 1, 2).merge().setValue("วิศวกรผู้ออกแบบ:").setHorizontalAlignment("right").setFontWeight("bold");
  sheet.getRange(34, 4, 1, 6).merge().setValue((pi.desName || "-") + "  (" + (pi.desPos || "-") + ")").setHorizontalAlignment("left");
  sheet.getRange(35, 4, 1, 6).merge().setValue("ใบอนุญาต กว.: " + (pi.desLic || "-")).setHorizontalAlignment("left");
  
  sheet.getRange(37, 2, 1, 2).merge().setValue("วิศวกรผู้ตรวจสอบ:").setHorizontalAlignment("right").setFontWeight("bold");
  sheet.getRange(37, 4, 1, 6).merge().setValue(pi.chkName || "-").setHorizontalAlignment("left");
  sheet.getRange(38, 4, 1, 6).merge().setValue("ใบอนุญาต กว.: " + (pi.chkLic || "-")).setHorizontalAlignment("left");
  
  sheet.getRange(45, 1, 1, 10).merge().setValue(pi.compName || "KSKY Engineering co., ltd.").setHorizontalAlignment("center").setFontSize(14).setFontWeight("bold").setFontColor("#1565c0");

  let r = 60; 
  
  let cb = sc.col_b * 100, ch = sc.col_h * 100;
  let bxb = sc.bmX_b * 100, bxh = sc.bmX_h * 100;
  let byb = sc.bmY_b * 100, byh = sc.bmY_h * 100;
  
  let Ic_X = 0.70 * ((ch * Math.pow(cb, 3)) / 12);
  let Ib_X = 0.35 * ((bxb * Math.pow(bxh, 3)) / 12);
  let Ic_Y = 0.70 * ((cb * Math.pow(ch, 3)) / 12);
  let Ib_Y = 0.35 * ((byb * Math.pow(byh, 3)) / 12);
  
  let Ec_Tcm2 = (15.1 * Math.sqrt(sc.fc)); 
  
  let MOT = 0;
  let max_delta_X_cm = 0, max_delta_Y_cm = 0;
  let sum_elastic_dxe_X_cm = 0, sum_elastic_dxe_Y_cm = 0;
  let max_th_X = 0, max_th_Y = 0;
  let typ_h_cm = 0;

  res.levels.forEach(l => {
    MOT += l.Fx * l.hx;
    if (l.hx > 0) {
      if (typ_h_cm === 0) typ_h_cm = l.hx * 100; 

      let dx_cm = l.dx_X / 10;
      let dy_cm = l.dx_Y / 10;
      if (dx_cm > max_delta_X_cm) max_delta_X_cm = dx_cm;
      if (dy_cm > max_delta_Y_cm) max_delta_Y_cm = dy_cm;

      sum_elastic_dxe_X_cm += (l.kx_X > 0) ? (l.Vx / l.kx_X) * 100 : 0;
      sum_elastic_dxe_Y_cm += (l.kx_Y > 0) ? (l.Vx / l.kx_Y) * 100 : 0;

      if (l.th_X > max_th_X) max_th_X = l.th_X;
      if (l.th_Y > max_th_Y) max_th_Y = l.th_Y;
    }
  });

  if(typ_h_cm === 0 && levels.length > 1) {
    typ_h_cm = (p.totalHeight / (levels.length - 1)) * 100;
  }
  
  let SF_X = (MOT > 0) ? (p.W_total * (p.totalSpanX / 2)) / MOT : 999;
  let SF_Y = (MOT > 0) ? (p.W_total * (p.totalSpanY / 2)) / MOT : 999;
  
  let H_allow_cm = (p.totalHeight * 100) / 500;
  let delta_a_cm = 0.01 * typ_h_cm;
  let max_delta_story_cm = Math.max(max_delta_X_cm, max_delta_Y_cm);
  let max_sum_elastic_cm = Math.max(sum_elastic_dxe_X_cm, sum_elastic_dxe_Y_cm);
  
  let Cd_val = parseFloat(p.Cd);
  if (isNaN(Cd_val) || Cd_val <= 0) Cd_val = 2.5; 
  let th_max = Math.min(0.5 / Cd_val, 0.25);
  
  let base_kx_Tcm = (res.levels[1] && res.levels[1].kx_X) ? res.levels[1].kx_X / 100 : 0;
  let base_ky_Tcm = (res.levels[1] && res.levels[1].kx_Y) ? res.levels[1].kx_Y / 100 : 0;

  let spanXArr = p.spanXStr.split(',').map(Number).filter(n => n > 0);
  let spanYArr = p.spanYStr.split(',').map(Number).filter(n => n > 0);
  let max_X = Math.max(...spanXArr);
  let max_Y = Math.max(...spanYArr);

  let span_S = Math.min(max_X, max_Y);
  let span_L = Math.max(max_X, max_Y);
  let m_ratio = span_S / span_L;
  
  let LL_short = 2 * ((gc.LL * span_S) / 3);
  let LL_long = 2 * ((gc.LL * span_S) / 3) * ((3 - Math.pow(m_ratio, 2)) / 2);
  let DL_short = gc.DL * span_S;
  let DL_long = gc.DL * span_L;
  
  let sumStart = r;
  sheet.getRange(r, 1, 1, 10).merge().setValue("SEISMIC ANALYSIS & STABILITY CHECKS").setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#e6f4ea"); r += 2;
  
  sheet.getRange(r, 1, 1, 10).merge().setValue("- ข้อมูลทั่วไปของอาคาร (Building Information)").setFontWeight("bold"); r++;
  sheet.getRange(r, 1).setValue("ความสูงรวม :").setHorizontalAlignment("right");
  sheet.getRange(r, 2, 1, 2).merge().setValue(sNum(p.totalHeight, 2)).setBackground("#f3f3f3").setHorizontalAlignment("center");
  sheet.getRange(r, 4).setValue("ม.").setHorizontalAlignment("left");
  sheet.getRange(r, 5).setValue("พื้นที่ต่อชั้น :").setHorizontalAlignment("right");
  sheet.getRange(r, 6).setValue(sNum(p.floorArea, 2)).setBackground("#f3f3f3").setHorizontalAlignment("center");
  sheet.getRange(r, 7).setValue("ตร.ม.").setHorizontalAlignment("left");
  sheet.getRange(r, 8).setValue("พื้นที่รวม :").setHorizontalAlignment("right");
  sheet.getRange(r, 9).setValue(sNum(p.totalArea, 2)).setBackground("#f3f3f3").setHorizontalAlignment("center");
  sheet.getRange(r, 10).setValue("ตร.ม.").setHorizontalAlignment("left");
  r += 2;

  sheet.getRange(r, 1, 1, 10).merge().setValue("- ข้อมูลที่ตั้งและพารามิเตอร์แผ่นดินไหว (Location & Seismic Parameters)").setFontWeight("bold"); r++;
  sheet.getRange(r, 1).setValue("จังหวัด :").setHorizontalAlignment("right");
  sheet.getRange(r, 2, 1, 2).merge().setValue(p.province || "-").setBackground("#f3f3f3").setHorizontalAlignment("center");
  sheet.getRange(r, 4).setValue("อำเภอ :").setHorizontalAlignment("right");
  sheet.getRange(r, 5, 1, 2).merge().setValue(p.amphoe || "-").setBackground("#f3f3f3").setHorizontalAlignment("center");
  sheet.getRange(r, 7).setValue("ชั้นดิน :").setHorizontalAlignment("right");
  sheet.getRange(r, 8).setValue(p.SiteClass || "-").setBackground("#f3f3f3").setHorizontalAlignment("center");
  sheet.getRange(r, 9).setValue("ความสำคัญ :").setHorizontalAlignment("right");
  sheet.getRange(r, 10).setValue(p.RiskCat || "-").setBackground("#f3f3f3").setHorizontalAlignment("center");
  r++;

  if (p.basinZone > 0) {
    sheet.getRange(r, 1).setValue("แอ่ง กทม. :").setHorizontalAlignment("right");
    sheet.getRange(r, 2, 1, 2).merge().setValue("Zone " + p.basinZone).setBackground("#f3f3f3").setHorizontalAlignment("center");
    sheet.getRange(r, 4).setValue("Sa(0.5s) :").setHorizontalAlignment("right");
    sheet.getRange(r, 5).setValue(sNum(p.Sa05, 3)).setBackground("#f3f3f3").setHorizontalAlignment("center");
    sheet.getRange(r, 6).setValue("g").setHorizontalAlignment("left"); 
  } else {
    sheet.getRange(r, 1).setValue("Ss :").setHorizontalAlignment("right");
    sheet.getRange(r, 2).setValue(sNum(p.Ss, 3)).setBackground("#f3f3f3").setHorizontalAlignment("center");
    sheet.getRange(r, 3).setValue("g").setHorizontalAlignment("left"); 
    sheet.getRange(r, 4).setValue("S1 :").setHorizontalAlignment("right");
    sheet.getRange(r, 5).setValue(sNum(p.S1, 3)).setBackground("#f3f3f3").setHorizontalAlignment("center");
    sheet.getRange(r, 6).setValue("g").setHorizontalAlignment("left"); 
    sheet.getRange(r, 7).setValue("Fa :").setHorizontalAlignment("right");
    sheet.getRange(r, 8).setValue(sNum(p.Fa, 3)).setBackground("#f3f3f3").setHorizontalAlignment("center");
    sheet.getRange(r, 9).setValue("Fv :").setHorizontalAlignment("right");
    sheet.getRange(r, 10).setValue(sNum(p.Fv, 3)).setBackground("#f3f3f3").setHorizontalAlignment("center");
  }
  r++;

  sheet.getRange(r, 1).setValue("ระบบรับแรง :").setHorizontalAlignment("right");
  sheet.getRange(r, 2, 1, 5).merge().setValue(p.systemType || "-").setBackground("#f3f3f3").setHorizontalAlignment("center");
  sheet.getRange(r, 7).setValue("R :").setHorizontalAlignment("right");
  sheet.getRange(r, 8).setValue(p.R).setBackground("#f3f3f3").setHorizontalAlignment("center");
  sheet.getRange(r, 9).setValue("Ie :").setHorizontalAlignment("right");
  sheet.getRange(r, 10).setValue(p.Ie).setBackground("#f3f3f3").setHorizontalAlignment("center");
  r++;

  sheet.getRange(r, 1).setValue("SDC :").setHorizontalAlignment("right");
  let sdcIndex = ["A","B","C","D"].indexOf(p.dc);
  let sdcThai = sdcIndex > -1 ? ["ก (A)","ข (B)","ค (C)","ง (D)"][sdcIndex] : p.dc;
  sheet.getRange(r, 2).setValue(sdcThai).setBackground("#f3f3f3").setHorizontalAlignment("center").setFontWeight("bold");
  sheet.getRange(r, 3).setValue("Cs :").setHorizontalAlignment("right");
  sheet.getRange(r, 4).setValue(sNum(p.Cs, 4)).setBackground("#f3f3f3").setHorizontalAlignment("center");
  sheet.getRange(r, 5).setValue("T :").setHorizontalAlignment("right");
  sheet.getRange(r, 6).setValue(sNum(p.T, 3)).setBackground("#f3f3f3").setHorizontalAlignment("center");
  sheet.getRange(r, 7).setValue("วินาที").setHorizontalAlignment("left"); 
  r += 2;

  sheet.getRange(r, 1, 1, 5).merge().setValue("แรงเฉือนที่ฐาน (Base Shear)");
  sheet.getRange(r, 6).setValue("V = CsW").setHorizontalAlignment("right");
  sheet.getRange(r, 7).setValue("'=").setHorizontalAlignment("center"); 
  sheet.getRange(r, 8).setValue(sNum(p.V, 3)).setHorizontalAlignment("center").setFontWeight("bold");
  sheet.getRange(r, 9).setValue("ตัน").setHorizontalAlignment("left").setFontWeight("bold");
  r += 2;
  
  sheet.getRange(r, 1, 1, 10).merge().setValue("- ตรวจสอบเสถียรภาพของอาคาร (Global Stability)").setFontWeight("bold"); r++;
  
  if (p.totalHeight > 25) {
      sheet.getRange(r, 1, 1, 10).merge().setValue("⚠️ หมายเหตุ: อาคารสูงเกิน 25 เมตร การใช้สมการ Simplified Stability อาจคลาดเคลื่อนสูง ควรใช้โปรแกรม 3D Finite Element").setFontColor("#d93025").setFontStyle("italic"); r++;
  }

  sheet.getRange(r, 1, 1, 4).merge().setValue("กรณีโครงสร้างคอนกรีตเสริมเหล็ก เมื่อ Ib=0.35Ig, Ic=0.70Ig");
  sheet.getRange(r, 6).setValue("fc' :").setHorizontalAlignment("right");
  sheet.getRange(r, 7).setValue(sc.fc).setHorizontalAlignment("center").setBackground("#f3f3f3");
  sheet.getRange(r, 8).setValue("ksc").setHorizontalAlignment("left"); r++;
  
  sheet.getRange(r, 1).setValue("ขนาดเสา (ซม.)").setHorizontalAlignment("right");
  sheet.getRange(r, 2).setValue("b :").setHorizontalAlignment("right");
  sheet.getRange(r, 3).setValue(cb).setBackground("#f3f3f3").setHorizontalAlignment("center");
  sheet.getRange(r, 4).setValue("h :").setHorizontalAlignment("right");
  sheet.getRange(r, 5).setValue(ch).setBackground("#f3f3f3").setHorizontalAlignment("center");
  sheet.getRange(r, 6).setValue("ขนาดคาน (ซม.)").setHorizontalAlignment("right");
  sheet.getRange(r, 7).setValue("[X] b :").setHorizontalAlignment("right");
  sheet.getRange(r, 8).setValue(bxb).setBackground("#f3f3f3").setHorizontalAlignment("center");
  sheet.getRange(r, 9).setValue("h :").setHorizontalAlignment("right");
  sheet.getRange(r, 10).setValue(bxh).setBackground("#f3f3f3").setHorizontalAlignment("center"); r++;
  
  sheet.getRange(r, 7).setValue("[Y] b :").setHorizontalAlignment("right");
  sheet.getRange(r, 8).setValue(byb).setBackground("#f3f3f3").setHorizontalAlignment("center");
  sheet.getRange(r, 9).setValue("h :").setHorizontalAlignment("right");
  sheet.getRange(r, 10).setValue(byh).setBackground("#f3f3f3").setHorizontalAlignment("center"); r += 2;
  
  sheet.getRange(r, 1, 1, 10).merge().setValue("ค่าโมเมนต์อินเนอร์เชียประสิทธิผล (cm^4) ของเสาและคานสำหรับแรงกระทำในทิศนั้นๆ"); r++;
  sheet.getRange(r, 2, 1, 2).merge().setValue("X-X (ทิศ E-W) :").setHorizontalAlignment("right");
  sheet.getRange(r, 4).setValue("Ic =").setHorizontalAlignment("right"); sheet.getRange(r, 5).setValue(sNum(Ic_X, 0)).setHorizontalAlignment("left");
  sheet.getRange(r, 6).setValue("Ib =").setHorizontalAlignment("right"); sheet.getRange(r, 7).setValue(sNum(Ib_X, 0)).setHorizontalAlignment("left"); r++;
  sheet.getRange(r, 2, 1, 2).merge().setValue("Y-Y (ทิศ N-S) :").setHorizontalAlignment("right");
  sheet.getRange(r, 4).setValue("Ic =").setHorizontalAlignment("right"); sheet.getRange(r, 5).setValue(sNum(Ic_Y, 0)).setHorizontalAlignment("left");
  sheet.getRange(r, 6).setValue("Ib =").setHorizontalAlignment("right"); sheet.getRange(r, 7).setValue(sNum(Ib_Y, 0)).setHorizontalAlignment("left"); r += 2;
  
  sheet.getRange(r, 1, 1, 4).merge().setValue("ค่าสติฟเนสของโครงสร้างชั้นฐาน (รับแรงตั้งฉากกับแนวแกน)");
  sheet.getRange(r, 6).setValue("Ec").setHorizontalAlignment("right");
  sheet.getRange(r, 7).setValue("'=").setHorizontalAlignment("center");
  sheet.getRange(r, 8).setValue(sNum(Ec_Tcm2, 2)).setHorizontalAlignment("center");
  sheet.getRange(r, 9).setValue("ตัน/ตร.ซม.").setHorizontalAlignment("left"); r++;
  
  sheet.getRange(r, 2).setValue("kx-x =").setHorizontalAlignment("right"); 
  sheet.getRange(r, 3).setValue(sNum(base_kx_Tcm, 2)).setHorizontalAlignment("center"); 
  sheet.getRange(r, 4).setValue("ตัน/ซม.").setHorizontalAlignment("left");
  
  sheet.getRange(r, 6).setValue("ky-y").setHorizontalAlignment("right"); 
  sheet.getRange(r, 7).setValue("'=").setHorizontalAlignment("center");
  sheet.getRange(r, 8).setValue(sNum(base_ky_Tcm, 2)).setHorizontalAlignment("center"); 
  sheet.getRange(r, 9).setValue("ตัน/ซม.").setHorizontalAlignment("left"); r += 2;
  
  sheet.getRange(r, 1, 1, 5).merge().setValue("ตรวจสอบความปลอดภัยต่อการพลิกคว่ำ (SF = W(L/2)/ΣFxhx >= 1.5)");
  sheet.getRange(r, 6).setValue("SF (X)").setHorizontalAlignment("right"); 
  sheet.getRange(r, 7).setValue("'=").setHorizontalAlignment("center");
  sheet.getRange(r, 8).setValue(sNum(SF_X, 2)).setHorizontalAlignment("center"); 
  setStatus(sheet.getRange(r, 10), SF_X >= 1.5 ? "PASS" : "FAIL"); r++; 
  
  sheet.getRange(r, 6).setValue("SF (Y)").setHorizontalAlignment("right"); 
  sheet.getRange(r, 7).setValue("'=").setHorizontalAlignment("center");
  sheet.getRange(r, 8).setValue(sNum(SF_Y, 2)).setHorizontalAlignment("center"); 
  setStatus(sheet.getRange(r, 10), SF_Y >= 1.5 ? "PASS" : "FAIL"); r += 2; 
  
  sheet.getRange(r, 1, 1, 4).merge().setValue("ค่าการเคลื่อนตัวต่อชั้นที่ยอมให้"); 
  sheet.getRange(r, 6).setValue("Δa = 0.01hx").setHorizontalAlignment("right"); 
  sheet.getRange(r, 7).setValue("'=").setHorizontalAlignment("center");
  sheet.getRange(r, 8).setValue(sNum(delta_a_cm, 2)).setHorizontalAlignment("center"); 
  sheet.getRange(r, 9).setValue("ซม.").setHorizontalAlignment("left"); r++;

  sheet.getRange(r, 1, 1, 4).merge().setValue("ค่าการเคลื่อนตัวสูงสุดที่ยอมให้"); 
  sheet.getRange(r, 6).setValue("Δmax = H/500").setHorizontalAlignment("right"); 
  sheet.getRange(r, 7).setValue("'=").setHorizontalAlignment("center"); 
  sheet.getRange(r, 8).setValue(sNum(H_allow_cm, 2)).setHorizontalAlignment("center");
  sheet.getRange(r, 9).setValue("ซม.").setHorizontalAlignment("left"); r++;
  
  sheet.getRange(r, 1, 1, 4).merge().setValue("ตรวจสอบการเคลื่อนตัวแต่ละชั้น"); 
  sheet.getRange(r, 6).setValue("Δa >= Δx").setHorizontalAlignment("right"); 
  sheet.getRange(r, 7).setValue("'=").setHorizontalAlignment("center");
  sheet.getRange(r, 8).setValue(p.totalHeight > 25 ? "N/A" : sNum(max_delta_story_cm, 2)).setHorizontalAlignment("center"); 
  setStatus(sheet.getRange(r, 10), p.totalHeight > 25 ? "-" : (max_delta_story_cm <= delta_a_cm ? "PASS" : "FAIL")); r++;
  
  sheet.getRange(r, 1, 1, 4).merge().setValue("ตรวจสอบการเคลื่อนตัวสูงสุด"); 
  sheet.getRange(r, 6).setValue("Δmax >= Σδxe").setHorizontalAlignment("right"); 
  sheet.getRange(r, 7).setValue("'=").setHorizontalAlignment("center");
  sheet.getRange(r, 8).setValue(p.totalHeight > 25 ? "N/A" : sNum(max_sum_elastic_cm, 2)).setHorizontalAlignment("center"); 
  setStatus(sheet.getRange(r, 10), p.totalHeight > 25 ? "-" : (max_sum_elastic_cm <= H_allow_cm ? "PASS" : "FAIL")); r++;
  
  sheet.getRange(r, 1, 1, 4).merge().setValue("ค่า สปส. เสถียรภาพสูงสุด"); 
  sheet.getRange(r, 6).setValue("θmax = 0.5/Cd").setHorizontalAlignment("right"); 
  sheet.getRange(r, 7).setValue("'=").setHorizontalAlignment("center"); 
  sheet.getRange(r, 8).setValue(sNum(th_max, 3)).setHorizontalAlignment("center"); r++;
  
  sheet.getRange(r, 1, 1, 4).merge().setValue("ตรวจสอบผลของโมเมนต์ลำดับที่สอง (0.1 < θ <= θmax)");
  sheet.getRange(r, 6).setValue("(X)").setHorizontalAlignment("right"); 
  sheet.getRange(r, 7).setValue("'=").setHorizontalAlignment("center");
  sheet.getRange(r, 8).setValue(p.totalHeight > 25 ? "N/A" : sNum(max_th_X, 4)).setHorizontalAlignment("center"); 
  let pStatX = "PASS"; if(p.totalHeight > 25) pStatX = "-"; else if (max_th_X > th_max) pStatX = "FAIL"; else if (max_th_X > 0.1) pStatX = "Amplify";
  setStatus(sheet.getRange(r, 10), pStatX); r++;
  
  sheet.getRange(r, 6).setValue("(Y)").setHorizontalAlignment("right"); 
  sheet.getRange(r, 7).setValue("'=").setHorizontalAlignment("center");
  sheet.getRange(r, 8).setValue(p.totalHeight > 25 ? "N/A" : sNum(max_th_Y, 4)).setHorizontalAlignment("center"); 
  let pStatY = "PASS"; if(p.totalHeight > 25) pStatY = "-"; else if (max_th_Y > th_max) pStatY = "FAIL"; else if (max_th_Y > 0.1) pStatY = "Amplify";
  setStatus(sheet.getRange(r, 10), pStatY); r += 2;

  sheet.getRange(r, 1, 1, 4).merge().setValue("- แรงในแนวดิ่งที่กระทำกับตัวอาคาร").setFontWeight("bold"); r++;
  sheet.getRange(r, 2, 1, 4).merge().setValue("แผ่นพื้นตัวแทน (หาจากช่วงเสากว้างสุด):");
  sheet.getRange(r, 6).setValue("S = " + span_S.toFixed(2)).setHorizontalAlignment("right");
  sheet.getRange(r, 7).setValue("L = " + span_L.toFixed(2)).setHorizontalAlignment("center");
  sheet.getRange(r, 8).setValue("ม.").setHorizontalAlignment("left"); r++;

  sheet.getRange(r, 2, 1, 3).merge().setValue("อัตราส่วนช่วงกว้างต่อช่วงยาว");
  sheet.getRange(r, 6).setValue("m = s/l").setHorizontalAlignment("right");
  sheet.getRange(r, 7).setValue("'=").setHorizontalAlignment("center");
  sheet.getRange(r, 8).setValue(sNum(m_ratio, 2)).setHorizontalAlignment("center"); r++;
  
  sheet.getRange(r, 2, 1, 3).merge().setValue("น้ำหนัก (LL) กระทำลงคานช่วงสั้น");
  sheet.getRange(r, 6).setValue("2(ws/3)").setHorizontalAlignment("right");
  sheet.getRange(r, 7).setValue("'=").setHorizontalAlignment("center");
  sheet.getRange(r, 8).setValue(sNum(LL_short, 2)).setHorizontalAlignment("center");
  sheet.getRange(r, 9).setValue("ตัน/ม.").setHorizontalAlignment("left"); r++;

  sheet.getRange(r, 2, 1, 3).merge().setValue("น้ำหนัก (LL) กระทำลงคานช่วงยาว");
  sheet.getRange(r, 6).setValue("2(ws/3)((3-m^2)/2)").setHorizontalAlignment("right");
  sheet.getRange(r, 7).setValue("'=").setHorizontalAlignment("center");
  sheet.getRange(r, 8).setValue(sNum(LL_long, 2)).setHorizontalAlignment("center");
  sheet.getRange(r, 9).setValue("ตัน/ม.").setHorizontalAlignment("left"); r++;
  
  sheet.getRange(r, 2, 1, 3).merge().setValue("น้ำหนัก (DL) กระทำลงคานด้านสั้น");
  sheet.getRange(r, 6).setValue("DL x S").setHorizontalAlignment("right");
  sheet.getRange(r, 7).setValue("'=").setHorizontalAlignment("center");
  sheet.getRange(r, 8).setValue(sNum(DL_short, 2)).setHorizontalAlignment("center");
  sheet.getRange(r, 9).setValue("ตัน/ม.").setHorizontalAlignment("left"); r++;
  
  sheet.getRange(r, 2, 1, 3).merge().setValue("น้ำหนัก (DL) กระทำลงคานด้านยาว");
  sheet.getRange(r, 6).setValue("DL x L").setHorizontalAlignment("right");
  sheet.getRange(r, 7).setValue("'=").setHorizontalAlignment("center");
  sheet.getRange(r, 8).setValue(sNum(DL_long, 2)).setHorizontalAlignment("center");
  sheet.getRange(r, 9).setValue("ตัน/ม.").setHorizontalAlignment("left"); r++;
  
  sheet.getRange(sumStart, 1, r - sumStart, 10).setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  let summaryEndRow = r - 1; 
  let tablesStartRow = 115; 

  let reportGapStart = summaryEndRow + 1;
  let reportGapEnd = tablesStartRow - 1; 

  if (reportGapStart <= reportGapEnd) {
      sheet.hideRows(reportGapStart, reportGapEnd - reportGapStart + 1);
      PropertiesService.getDocumentProperties().setProperty('reportGapStart', reportGapStart.toString());
      PropertiesService.getDocumentProperties().setProperty('reportGapEnd', reportGapEnd.toString());
  }

  r = tablesStartRow; 
  
  sheet.getRange(r, 1).setValue("ตารางแจกแจงรายชั้น (Detailed Story Tables)").setFontSize(12).setFontWeight("bold"); r += 2;
  
  sheet.getRange(r, 1).setValue("1. VERTICAL DISTRIBUTION OF FORCES").setFontWeight("bold"); r++;
  let head1 = [["Story", "Height hx", "Weight wx", "wx * hx^k", "Cvx", "Force Fx", "Shear Vx"]];
  sheet.getRange(r, 1, 1, 7).setValues(head1).setBackground(CONFIG.colors.tableHeader).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setWrap(true).setVerticalAlignment("middle"); r++;
  
  let table1 = [];
  for (let i = 0; i < levels.length; i++) {
    let l = levels[i]; let cvx = (p.V > 0) ? (l.Fx / p.V) : 0;
    table1.push([ l.name, sNum(l.hx, 2), sNum(l.wx, 2), sNum(l.wx_hx_k, 1), sNum(cvx, 4), sNum(l.Fx, 3), l.Vx ? sNum(l.Vx, 3) : "-" ]);
  }
  sheet.getRange(r, 1, table1.length, 7).setValues(table1).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment("center").setWrap(true).setVerticalAlignment("middle");
  for (let i = 0; i < table1.length; i++) if (i % 2 !== 0) sheet.getRange(r + i, 1, 1, 7).setBackground(CONFIG.colors.tableRowEven);
  r += table1.length + 2;

  function writeStabilityTable(title, isX) {
    sheet.getRange(r, 1).setValue(title).setFontWeight("bold"); r++;
    let head2 = [["Story", "kx (T/m)", "Px (T)", "Δx (mm)", "Drift Ratio", "Drift Stat", "P-Delta θ", "P-Delta Stat", "Overturn Mx"]];
    sheet.getRange(r, 1, 1, 9).setValues(head2).setBackground("#00695c").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setWrap(true).setVerticalAlignment("middle"); r++;
    let table2 = [];
    for (let i = 0; i < levels.length; i++) {
      let l = levels[i];
      if (l.hx > 0) { table2.push([ l.name, isX ? sNum(l.kx_X, 0) : sNum(l.kx_Y, 0), sNum(l.Px, 1), p.totalHeight > 25 ? "N/A" : (isX ? sNum(l.dx_X, 2) : sNum(l.dx_Y, 2)), p.totalHeight > 25 ? "N/A" : (isX ? sNum(l.drift_X, 5) : sNum(l.drift_Y, 5)), p.totalHeight > 25 ? "-" : (isX ? (l.dStat_X || "-") : (l.dStat_Y || "-")), p.totalHeight > 25 ? "N/A" : (isX ? sNum(l.th_X, 4) : sNum(l.th_Y, 4)), p.totalHeight > 25 ? "-" : (isX ? (l.pStat_X || "-") : (l.pStat_Y || "-")), sNum(l.Mx, 2) ]); }
    }
    if (table2.length > 0) {
      sheet.getRange(r, 1, table2.length, 9).setValues(table2).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment("center").setWrap(true).setVerticalAlignment("middle");
      for (let i = 0; i < table2.length; i++) {
        if (i % 2 !== 0) sheet.getRange(r + i, 1, 1, 9).setBackground(CONFIG.colors.tableRowEven);
        let dStat = table2[i][5]; let pStat = table2[i][7];
        if (dStat === "FAIL") sheet.getRange(r + i, 6).setBackground("#fce8e6").setFontColor("#c5221f").setFontWeight("bold"); else if (dStat === "OK") sheet.getRange(r + i, 6).setFontColor("#137333").setFontWeight("bold");
        if (pStat && pStat.includes("FAIL")) sheet.getRange(r + i, 8).setBackground("#fce8e6").setFontColor("#c5221f").setFontWeight("bold"); else if (pStat && pStat.includes("Amplify")) sheet.getRange(r + i, 8).setBackground("#fef7e0").setFontColor("#b06000").setFontWeight("bold"); else if (pStat === "OK") sheet.getRange(r + i, 8).setFontColor("#137333").setFontWeight("bold");
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
  sheet.getRange(r, 1, 1, 7).setValues(head3).setBackground("#455a64").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setWrap(true).setVerticalAlignment("middle"); r++;
  let table3 = [];
  for (let i = 0; i < levels.length; i++) { if (levels[i].hx > 0) { let f = levels[i].Fx || 0; table3.push([ levels[i].name, sNum(f, 3), sNum(f, 3), sNum(1.3 * f, 3), sNum(0.7 * f, 3), sNum(1.3 * f, 3), sNum(0.7 * f, 3) ]); } }
  if (table3.length > 0) { 
    sheet.getRange(r, 1, table3.length, 7).setValues(table3).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment("center").setWrap(true).setVerticalAlignment("middle"); 
    for (let i = 0; i < table3.length; i++) if (i % 2 !== 0) sheet.getRange(r + i, 1, 1, 7).setBackground(CONFIG.colors.tableRowEven); 
  }
}

// --- DRAWING LOGIC (DUAL-AXIS PLAN) ---
function generateBlueprintFromData(rawSpanX, rawHeight, rawLoads, rawSpanY, rawPointLoads, calcResult) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let planSheet = ss.getSheetByName(CONFIG.sheetPlan);
  if (!planSheet) planSheet = ss.insertSheet(CONFIG.sheetPlan);
  
  try { 
      var maxR = planSheet.getMaxRows();
      var maxC = planSheet.getMaxColumns();
      if(maxR > 0 && maxC > 0) {
          planSheet.showRows(1, maxR);
          planSheet.showColumns(1, maxC);
          planSheet.getRange(1, 1, maxR, maxC).breakApart(); 
      }
  } catch(e) {}
  planSheet.clear();

  if (planSheet.getMaxColumns() < 135) {
    planSheet.insertColumnsAfter(planSheet.getMaxColumns(), 135 - planSheet.getMaxColumns());
  }

  const spansX_meters = rawSpanX.split(',').map(Number);
  const spansY_meters = rawSpanY.split(',').map(Number);
  const heights_meters = rawHeight.split(',').map(Number).filter(n => n > 0); 
  const loads_val = rawLoads.split(',').map(Number);
  const ptLoads_val = rawPointLoads.split(',').map(Number);
  const pi = (calcResult && calcResult.params && calcResult.params.projectInfo) ? calcResult.params.projectInfo : {};

  let totalWidthX_m = spansX_meters.reduce((sum, val) => sum + (val || 0), 0);
  let totalWidthY_m = spansY_meters.reduce((sum, val) => sum + (val || 0), 0);
  let totalHeight_m = heights_meters.reduce((sum, val) => sum + (val || 0), 0);
  let max_span_m = Math.max(totalWidthX_m, totalWidthY_m);

  let res = Math.max(max_span_m / 40, totalHeight_m / 50);
  if (res < 0.1) res = 0.1;

  const spansX_cells = spansX_meters.map(n => Math.max(6, Math.round((n || 0) / res)));
  const spansY_cells = spansY_meters.map(n => Math.max(6, Math.round((n || 0) / res)));
  const heights_cells = heights_meters.map(n => Math.max(6, Math.round((n || 0) / res)));

  let totalWidthX_cells = spansX_cells.reduce((sum, val) => sum + val, 0);
  let totalWidthY_cells = spansY_cells.reduce((sum, val) => sum + val, 0);
  let totalHeight_cells = heights_cells.reduce((sum, val) => sum + val, 0);

  let maxF = 0; for (let i = 0; i < ptLoads_val.length; i++) if (ptLoads_val[i] > maxF) maxF = ptLoads_val[i];
  let maxV = (calcResult && calcResult.params && calcResult.params.V) ? calcResult.params.V : 0;

  let maxForce = Math.max(maxF, maxV, 1);
  let dScale = 12 / maxForce; 

  // ปรับจาก 45 เป็น 36 เพื่อขยับอาคารให้ชิดเส้น Dimension เข้ามาอีก 9 ช่อง
  const requiredLeftSpace = 36; 
  
  const totalRowsNeeded = 800; 

  if (totalRowsNeeded > planSheet.getMaxRows()) planSheet.insertRowsAfter(planSheet.getMaxRows(), totalRowsNeeded - planSheet.getMaxRows());

  planSheet.setColumnWidths(1, planSheet.getMaxColumns(), CONFIG.cellSizePx);
  planSheet.setRowHeights(1, totalRowsNeeded, CONFIG.cellSizePx);
  
  planSheet.getRange(1, 1, totalRowsNeeded, 130).setBackground("#ffffff");

  let infoCol = requiredLeftSpace;
  
  let planRightCol = requiredLeftSpace + Math.max(totalWidthX_cells, totalWidthY_cells) - 10;
  if(planRightCol < requiredLeftSpace + 40) planRightCol = requiredLeftSpace + 40;
  planSheet.getRange(2, planRightCol, 1, 15).merge().setValue("Project: " + (pi.projName || "-")).setHorizontalAlignment("right").setFontWeight("bold").setFontColor("#555");
  planSheet.getRange(3, planRightCol, 1, 15).merge().setValue("Date: " + fmtDate(pi.projDate)).setHorizontalAlignment("right").setFontColor("#555");
  planSheet.getRange(4, planRightCol, 1, 15).merge().setValue("Designed By: " + (pi.desName || "-") + " (กว. " + (pi.desLic || "-") + ")").setHorizontalAlignment("right").setFontColor("#555");

  if (calcResult && calcResult.params) {
    planSheet.getRange(2, infoCol, 1, 30).merge().setValue("BUILDING SUMMARY").setFontWeight("bold").setFontSize(10).setFontColor("#1565c0").setHorizontalAlignment("left");
    planSheet.getRange(3, infoCol, 1, 30).merge().setValue("• Floor Area: " + sNum(calcResult.params.floorArea, 2) + " sq.m/floor").setFontSize(9).setFontWeight("bold").setHorizontalAlignment("left");
    planSheet.getRange(4, infoCol, 1, 30).merge().setValue("• Total Area: " + sNum(calcResult.params.totalArea, 2) + " sq.m").setFontSize(9).setFontWeight("bold").setHorizontalAlignment("left");
  }

  function drawSet(startRow, spansHoriz_m, spansHoriz_c, spansDepth_m, spansDepth_c, axisName, loadPrefix) {
    let drawingWidth = 0; for (let i = 0; i < spansHoriz_c.length; i++) drawingWidth += spansHoriz_c[i];
    let startCol = requiredLeftSpace;
    let currentRow = startRow;
    
    planSheet.getRange(currentRow - 6, startCol).setValue("ELEVATION VIEW (" + axisName + " Axis)").setFontSize(12).setFontWeight("bold").setFontColor("#1565c0");
    
    if (calcResult && calcResult.params) {
        let dimRange = planSheet.getRange(currentRow, 4, totalHeight_cells, 2);
        dimRange.merge()
          .setValue("Total Height = " + sNum(calcResult.params.totalHeight, 2) + " m")
          .setFontColor(CONFIG.colors.dimText)
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle")
          .setFontWeight("bold")
          .setFontSize(9)
          .setTextRotation(90)
          .setBorder(true, true, true, false, null, null, CONFIG.colors.dimText, SpreadsheetApp.BorderStyle.SOLID);
    }
    
    for (let idx = 0; idx < heights_cells.length; idx++) {
      let currentX = startCol;
      let h_cells = heights_cells[idx];
      let h_meters = heights_meters[idx] || 0;
      
      let pLoadIndex = (ptLoads_val.length - 1) - idx;
      let dLoadIndex = (loads_val.length - 1) - idx;
      let pLoad = ptLoads_val[pLoadIndex] || 0;
      let dLoad = loads_val[dLoadIndex] || 0;
      
      let dimHeight = h_cells - 2; 
      if (dimHeight < 1) dimHeight = 1;
      planSheet.getRange(currentRow + 1, startCol - 4, dimHeight, 3).merge().setValue(h_meters.toFixed(2) + "m").setFontColor(CONFIG.colors.dimText).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(8);
      
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

    let columnX = startCol; let colPositions = [startCol];
    for (let i = 0; i < spansHoriz_c.length; i++) { columnX += spansHoriz_c[i]; colPositions.push(columnX); }
    for (let i = 0; i < colPositions.length; i++) {
      drawColumnStump(planSheet, currentRow, colPositions[i], CONFIG.stumpHeight);
      drawFixedSupport(planSheet, currentRow + CONFIG.stumpHeight, colPositions[i]);
    }

    let gridX = startCol; let gridNum = 1; let labelRow = currentRow + CONFIG.stumpHeight + 4;
    createLabelBox(planSheet, labelRow, gridX - 1, gridNum, CONFIG.colors.gridLabel); gridNum++;
    
    for (let i = 0; i < spansHoriz_c.length; i++) {
      let spaceWidth = spansHoriz_c[i] - 2; 
      if (spaceWidth > 0) { planSheet.getRange(labelRow, gridX + 1, 2, spaceWidth).merge().setValue((spansHoriz_m[i] || 0).toFixed(2) + "m").setFontColor(CONFIG.colors.dimText).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(8); }
      gridX += spansHoriz_c[i]; createLabelBox(planSheet, labelRow, gridX - 1, gridNum, CONFIG.colors.gridLabel); gridNum++;
    }
    
    let totalSpanM = 0; for (let i = 0; i < spansHoriz_m.length; i++) totalSpanM += (spansHoriz_m[i] || 0);
    let totLabelRow = labelRow + 3;
    planSheet.getRange(totLabelRow, startCol, 2, drawingWidth).merge()
      .setValue("Total Width = " + sNum(totalSpanM, 2) + " m")
      .setFontColor(CONFIG.colors.dimText)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setFontWeight("bold")
      .setFontSize(9)
      .setBorder(false, true, true, true, null, null, CONFIG.colors.dimText, SpreadsheetApp.BorderStyle.SOLID); 

    currentRow += 16;
    planSheet.getRange(currentRow - 4, startCol).setValue("PLAN VIEW (" + axisName + " Axis)").setFontSize(12).setFontWeight("bold").setFontColor("#1565c0");
    
    gridX = startCol; gridNum = 1;
    createLabelBox(planSheet, currentRow - 3, gridX - 1, gridNum, CONFIG.colors.gridLabel); gridNum++;
    
    for (let i = 0; i < spansHoriz_c.length; i++) {
      let spaceWidth = spansHoriz_c[i] - 2;
      if (spaceWidth > 0) { planSheet.getRange(currentRow - 3, gridX + 1, 2, spaceWidth).merge().setValue((spansHoriz_m[i] || 0).toFixed(2) + "m").setFontColor(CONFIG.colors.dimText).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(8); }
      gridX += spansHoriz_c[i]; createLabelBox(planSheet, currentRow - 3, gridX - 1, gridNum, CONFIG.colors.gridLabel); gridNum++;
    }
    
    for (let idx = 0; idx < spansDepth_c.length; idx++) {
      let currentX = startCol; let hc = spansDepth_c[idx];
      let charCode = String.fromCharCode(65 + idx);
      
      createLabelBox(planSheet, currentRow - 1, startCol - 3, charCode, CONFIG.colors.gridLabel);
      let dimHeight = hc - 2; if (dimHeight < 1) dimHeight = 1;
      planSheet.getRange(currentRow + 1, startCol - 4, dimHeight, 3).merge().setValue((spansDepth_m[idx] || 0).toFixed(2) + "m").setFontColor(CONFIG.colors.dimText).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(8);
      
      for (let i = 0; i < spansHoriz_c.length; i++) {
        let room = planSheet.getRange(currentRow, currentX, hc, spansHoriz_c[i]);
        room.setBackground(CONFIG.colors.fillTop).setBorder(true, true, true, true, null, null, "#90a4ae", SpreadsheetApp.BorderStyle.SOLID);
        currentX += spansHoriz_c[i];
      }
      if (idx === spansDepth_c.length - 1) { createLabelBox(planSheet, currentRow + hc - 1, startCol - 3, String.fromCharCode(65 + idx + 1), CONFIG.colors.gridLabel); }
      currentRow += hc;
    }
    return currentRow + 10;
  }

  let nextRow = drawSet(12, spansX_meters, spansX_cells, spansY_meters, spansY_cells, "E-W / X", "Fx");
  
  let yAxisStartRow = 240; 
  while (yAxisStartRow <= nextRow + 20) {
      yAxisStartRow += 240; 
  }
  
  let gapStart = nextRow + 1;
  let gapEnd = yAxisStartRow - 10; 
  
  if (gapStart < gapEnd) {
      planSheet.hideRows(gapStart, gapEnd - gapStart + 1);
      PropertiesService.getDocumentProperties().setProperty('planGapStart', gapStart.toString());
      PropertiesService.getDocumentProperties().setProperty('planGapEnd', gapEnd.toString());
  }

  drawSet(yAxisStartRow, spansY_meters, spansY_cells, spansX_meters, spansX_cells, "N-S / Y", "Fy");
  
  if (calcResult && calcResult.params && calcResult.params.stiffnessConfig) {
    const sc = calcResult.params.stiffnessConfig;
    let secCol = 131; 
    
    planSheet.getRange(4, secCol, 1, 10).merge().setValue("SECTION DETAILS").setFontWeight("bold").setFontSize(10).setFontColor("#1565c0").setHorizontalAlignment("left");
    planSheet.getRange(6, secCol, 1, 10).merge().setValue("• Column: " + sNum(sc.col_b, 2) + " x " + sNum(sc.col_h, 2) + " m.").setFontSize(9).setFontWeight("bold").setHorizontalAlignment("left");
    planSheet.getRange(8, secCol, 1, 10).merge().setValue("• Beam X: " + sNum(sc.bmX_b, 2) + " x " + sNum(sc.bmX_h, 2) + " m.").setFontSize(9).setFontWeight("bold").setHorizontalAlignment("left");
    planSheet.getRange(10, secCol, 1, 10).merge().setValue("• Beam Y: " + sNum(sc.bmY_b, 2) + " x " + sNum(sc.bmY_h, 2) + " m.").setFontSize(9).setFontWeight("bold").setHorizontalAlignment("left");
  }
  planSheet.setHiddenGridlines(true);
}

function drawShearForceRight(sheet, row, col, height, val, scale) {
  let arrowLength = Math.max(3, Math.ceil(val * scale));
  if (arrowLength > 12) arrowLength = 12; 
  let line = ""; for (let i = 0; i < Math.max(1, arrowLength); i++) line += "─";
  let text = "←" + line + " V = " + (val || 0).toFixed(2) + "T";
  sheet.getRange(row + Math.floor(height / 2) - 1, col, 2, arrowLength + 16).merge().setValue(text).setHorizontalAlignment("left").setVerticalAlignment("middle").setFontColor(CONFIG.colors.shearArrow).setFontWeight("bold").setFontSize(9);
}

function drawLateralLoad(sheet, beamRow, startCol, val, scale, prefix) {
  let arrowLength = Math.max(3, Math.ceil(val * scale));
  if (arrowLength > 12) arrowLength = 12; 
  let rightGap = 1; let textSpace = 16; 
  let targetCol = startCol - arrowLength - rightGap - textSpace;
  if (targetCol < 1) targetCol = 1; 
  let line = ""; for (let i = 0; i < Math.max(1, arrowLength); i++) line += "─";
  let valStr = (typeof val === 'number') ? val.toFixed(3) : val;
  let text = prefix + " = " + valStr + "T " + line + "→";
  sheet.getRange(beamRow - 1, targetCol, 2, arrowLength + textSpace).merge().setValue(text).setHorizontalAlignment("right").setVerticalAlignment("middle").setFontColor(CONFIG.colors.loadArrow).setFontWeight("bold").setFontSize(9);
}

function drawLoadArrows(sheet, beamRow, startCol, width) {
  if (beamRow - 1 > 0) {
    let arrows = ""; let count = Math.max(1, Math.floor(width - 1));
    for (let i = 0; i < count; i++) arrows += "↓ ";
    sheet.getRange(beamRow - 1, startCol, 1, width).merge().setValue(arrows).setHorizontalAlignment("center").setVerticalAlignment("bottom").setFontColor(CONFIG.colors.loadArrow).setFontSize(8).setFontWeight("bold");
  }
}

function drawLoadLabel(sheet, beamRow, startCol, width, val) {
  if (beamRow - 3 > 0) { sheet.getRange(beamRow - 3, startCol, 2, width).merge().setValue(val + " T/m²").setHorizontalAlignment("center").setVerticalAlignment("middle").setFontColor(CONFIG.colors.loadText).setFontSize(9).setFontWeight("bold"); }
}

function createLabelBox(sheet, row, col, text, color, align, isBold, fontSize) {
  if (row < 1 || col < 1) return;
  if (align === undefined) align = "center"; if (isBold === undefined) isBold = true; if (fontSize === undefined) fontSize = 8;
  let range = sheet.getRange(row, col, 2, 2);
  range.merge().setValue(text).setFontColor(color).setBackground(CONFIG.colors.labelBg).setBorder(true, true, true, true, null, null, "#dddddd", SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment(align).setVerticalAlignment("middle").setFontSize(fontSize);
  if (isBold) range.setFontWeight("bold");
}

function drawFixedSupport(sheet, row, centerX) {
  let startX = centerX - 1; 
  if (startX > 0) {
    sheet.getRange(row, startX, 2, 3).merge().setBackground(CONFIG.colors.support).setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.getRange(row + 2, startX - 1, 1, 5).merge().setValue("/ / / / /").setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(8).setFontColor("#757575").setFontStyle("italic");
  }
}

function drawColumnStump(sheet, row, x, height) {
  sheet.getRange(row, x, height, 1).setBorder(null, true, null, null, null, null, CONFIG.colors.beam, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}