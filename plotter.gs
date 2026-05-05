/**
 * Generates a 3-chart dashboard on a separate sheet with dynamic 25% padding.
 * Automatically handles dual-axes and prevents lines from touching boundaries.
 */
function updateThreeGraphDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getActiveSheet(); 
  const dashName = "Dashboard";
  
  // 1. Setup Dashboard Sheet
  let dashSheet = ss.getSheetByName(dashName);
  if (!dashSheet) {
    dashSheet = ss.insertSheet(dashName);
  }
  dashSheet.setHiddenGridlines(true); // Whiteboard background

  const lastRow = dataSheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log("No data found to plot.");
    return;
  }

  // 2. Data Collection & Buffer Calculation
  const goldVals = dataSheet.getRange("C2:C" + lastRow).getValues().flat().filter(Number);
  const silverVals = dataSheet.getRange("D2:D" + lastRow).getValues().flat().filter(Number);
  const ratioVals = dataSheet.getRange("G2:G" + lastRow).getValues().flat().filter(Number);

  // Function to calculate a 25% padding for the view window
  const getViewWindow = (vals) => {
    const min = Math.min(...vals);
    const max = Math.max(...vals);
    const range = max - min;
    // 25% buffer of the data spread
    const padding = range * 0.25 || min * 0.05; 
    return { min: min - padding, max: max + padding };
  };

  const gW = getViewWindow(goldVals);
  const sW = getViewWindow(silverVals);
  const rW = getViewWindow(ratioVals);

  // 3. Define Ranges
  const rangeA = dataSheet.getRange("A1:A" + lastRow); // Datetime
  const rangeC = dataSheet.getRange("C1:C" + lastRow); // Gold
  const rangeD = dataSheet.getRange("D1:D" + lastRow); // Silver
  const rangeG = dataSheet.getRange("G1:G" + lastRow); // Ratio

  // 4. Clean Existing Charts
  const existingCharts = dashSheet.getCharts();
  existingCharts.forEach(c => dashSheet.removeChart(c));

  // --- CHART A: Gold (Primary) vs Silver (Secondary) ---
  const chartA = dashSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(rangeA).addRange(rangeC).addRange(rangeD)
    .setPosition(2, 2, 0, 0)
    .setOption('title', 'A: Gold Rate (Left) vs Silver Rate (Right)')
    .setOption('height', 400)
    .setOption('width', 800)
    .setOption('series', {
      0: {targetAxisIndex: 0, color: '#f1c232', lineWidth: 3}, // Gold
      1: {targetAxisIndex: 1, color: '#999999', lineWidth: 3}  // Silver
    })
    .setOption('vAxes', {
      0: {title: 'Gold Price', viewWindow: {min: gW.min, max: gW.max}},
      1: {title: 'Silver Price', viewWindow: {min: sW.min, max: sW.max}}
    })
    .build();

  // --- CHART B: Ratio (Primary) vs Silver (Secondary) ---
  const chartB = dashSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(rangeA).addRange(rangeG).addRange(rangeD)
    .setPosition(23, 2, 0, 0)
    .setOption('title', 'B: Gold/Silver Ratio vs Silver Price')
    .setOption('height', 400)
    .setOption('width', 800)
    .setOption('series', {
      0: {targetAxisIndex: 0, color: '#cc0000', lineWidth: 3}, // Ratio
      1: {targetAxisIndex: 1, color: '#999999', lineWidth: 3}  // Silver
    })
    .setOption('vAxes', {
      0: {title: 'Ratio', viewWindow: {min: rW.min, max: rW.max}},
      1: {title: 'Silver Price', viewWindow: {min: sW.min, max: sW.max}}
    })
    .build();

  // --- CHART C: Ratio (Primary) vs Gold (Secondary) ---
  const chartC = dashSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(rangeA).addRange(rangeG).addRange(rangeC)
    .setPosition(44, 2, 0, 0)
    .setOption('title', 'C: Ratio vs Gold Price')
    .setOption('height', 400)
    .setOption('width', 800)
    .setOption('series', {
      0: {targetAxisIndex: 0, color: '#cc0000', lineWidth: 3}, // Ratio
      1: {targetAxisIndex: 1, color: '#f1c232', lineWidth: 3}  // Gold
    })
    .setOption('vAxes', {
      0: {title: 'Ratio', viewWindow: {min: rW.min, max: rW.max}},
      1: {title: 'Gold Price', viewWindow: {min: gW.min, max: gW.max}}
    })
    .build();

  // 5. Insert Charts into Dashboard
  dashSheet.insertChart(chartA);
  dashSheet.insertChart(chartB);
  dashSheet.insertChart(chartC);
  
  ss.setActiveSheet(dashSheet);
}
