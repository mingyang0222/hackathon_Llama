// add data code (done)
function showInputForm() {
    var html = HtmlService.createHtmlOutputFromFile('InputForm')
        .setTitle('Enter Data');
    SpreadsheetApp.getUi().showSidebar(html);
  }
  
  function processForm(formObject) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Format the date and currency values
    var date = new Date(formObject.date);
    var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    var sales = parseFloat(formObject.sales).toFixed(2);
    var purchases = parseFloat(formObject.purchases).toFixed(2);
    var rental = parseFloat(formObject.rental).toFixed(2);
    var transportation = parseFloat(formObject.transportation).toFixed(2);
    var others1 = parseFloat(formObject.others1).toFixed(2);
    var others2 = parseFloat(formObject.others2).toFixed(2);
    
    // Append the data to the sheet with the formatted values
    var newRow = [
      formattedDate,
      `$${sales}`,
      `$${purchases}`,
      `$${rental}`,
      `$${transportation}`,
      `$${others1}`,
      `$${others2}`
    ];
    
    sheet.appendRow(newRow);
    
    // Align the added data to the right side
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    var range = sheet.getRange(lastRow, 1, 1, lastColumn);
    range.setHorizontalAlignment('right');
    
    // Show a success message to the user
    SpreadsheetApp.getUi().alert('Data successfully added!');
  }
  
  // append data 
  function showMonthYearForm() {
    var html = HtmlService.createHtmlOutputFromFile('MonthYearForm')
        .setTitle('Calculate Statistics');
    SpreadsheetApp.getUi().showSidebar(html);
  }
  
  function createNewSheet(sheetName) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    
    if (sheet) {
      Logger.log('Sheet "%s" already exists.', sheetName);
    } else {
      sheet = ss.insertSheet(sheetName);
      Logger.log('Sheet "%s" has been created.', sheetName);
  
      var header = ['Month', 'Year', 'Sales', 'Purchases', 'Rental', 'Transportation', 'Others1', 'Others2', 'Time' , 'Total cost', 'Profit'];
      sheet.appendRow(header);
    }
  }
  
  function calculateAndAppendSales(month, year) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheetByName('SQL_Data');
    var statSheet = ss.getSheetByName('Statistic') || createNewSheet('Statistic');
  
    var data = dataSheet.getDataRange().getValues();
    var [salesTotal, purchasesTotal, rentalTotal, transportationTotal, others1Total, others2Total] = [0, 0, 0, 0, 0, 0];
  
    if (Array.isArray(data) && data.length > 1) {
      for (let i = 1; i < data.length; i++) {
          let row = data[i];
          let date = new Date(row[0]);
  
          if (!isNaN(date) && date.getMonth() + 1 === month && date.getFullYear() === year) {
              salesTotal += addIfNotNaN(parseFloat(row[1]));
              purchasesTotal += addIfNotNaN(parseFloat(row[2]));
              rentalTotal += addIfNotNaN(parseFloat(row[3]));
              transportationTotal += addIfNotNaN(parseFloat(row[4]));
              others1Total += addIfNotNaN(parseFloat(row[5]));
              others2Total += addIfNotNaN(parseFloat(row[6]));
          }
      }
  }
  
    var num_rows = statSheet.getLastRow();
    var totalCost = purchasesTotal + rentalTotal + transportationTotal + others1Total + others2Total;
    var profit = salesTotal - totalCost;
    var newRow = [month, year, salesTotal, purchasesTotal, rentalTotal, transportationTotal, others1Total, others2Total, num_rows , totalCost, profit];
    statSheet.appendRow(newRow);
  }
  
  function addIfNotNaN(value) {
    return isNaN(value) ? 0 : value;
  }
  
  function processMonthYearForm(formObject) {
    var month = parseInt(formObject.month);
    var year = parseInt(formObject.year);
    calculateAndAppendSales(month, year);
  }
  
  //predict
  
  function calculatePredictions() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Statistic'); // Replace with your actual sheet name
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();
  
    // Extract sales and purchases data
    var salesData = data.slice(1).map(function(row) { return row[2]; }); // Assuming sales is in the 3rd column
    var totalCostData = data.slice(1).map(function(row) { return row[9]; }); // Assuming purchases is in the 10th column
    var times = data.slice(1).map(function(row) { return row[8]; }); // Assuming data starts from row 2
  
    // Calculate slope and intercept for sales
    var slopeSales = getSlope(times, salesData);
    var interceptSales = getIntercept(times, salesData, slopeSales);
  
    // Calculate slope and intercept for purchases
    var slopeTotalCost = getSlope(times, totalCostData);
    var interceptTotalCost = getIntercept(times, totalCostData, slopeTotalCost);
  
    // Get the last month and year values
    var lastRowMonth = data[data.length - 1][0];
    var lastRowYear = data[data.length - 1][1];
    var lastRowTimes = data[data.length - 1][8];
  
    // Calculate the next month and year
    var nextMonth = lastRowMonth + 1;
    var nextYear = lastRowYear;
    if (nextMonth === 13) {
      nextMonth = 1;
      nextYear += 1;
    }
  
    // Predict sales and purchases for the next month
    var predictedRow = sheet.getLastRow() ;
    var predictedSales = slopeSales * (lastRowTimes+1) + interceptSales;
    var predictedTotalCost = slopeTotalCost * (lastRowTimes+1) + interceptTotalCost;
    var predictedProfit = predictedSales - predictedTotalCost;
  
    // Create a new row with predicted values
    var newRow = [nextMonth, nextYear, predictedSales, '', '', '', '', '', predictedRow, predictedTotalCost, predictedProfit];
    sheet.appendRow(newRow);
  
    highlightLastRowYellow(sheet)
  }
  
  // Helper function to calculate slope
  function getSlope(x, y) {
    var n = x.length;
    var sumX = 0, sumY = 0, sumXY = 0, sumXX = 0;
    for (var i = 0; i < n; i++) {
      sumX += x[i];
      sumY += y[i];
      sumXY += x[i] * y[i];
      sumXX += x[i] * x[i];
    }
    return ((n * sumXY - sumX * sumY) / (n * sumXX - sumX * sumX));
  }
  
  // Helper function to calculate intercept
  function getIntercept(x, y, slope) {
    var n = x.length;
    var sumX = 0, sumY = 0;
    for (var i = 0; i < n; i++) {
      sumX += x[i];
      sumY += y[i];
    }
    return (sumY - slope * sumX) / n;
  }
  
  function highlightLastRowYellow(sheet) {
    // Get the index of the last row
    const lastRow = sheet.getLastRow();
    
    // Check if there are any rows in the sheet
    if (lastRow > 0) {
      // Get the range of the last row (all columns)
      const range = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());
      
      // Set the background color to yellow
      range.setBackground('#FFFF00'); // Yellow color
    } else {
      Logger.log('The sheet is empty.');
    }
  }
  
  //create bar chart
  function createBarChartForMonthlyData() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Statistic');
    
    // Find the last row with data
    const lastRow = sheet.getLastRow();
    
    // Define the range for the chart
    // Adjust the columns if they are different in your sheet
    const range = sheet.getRange(`A1:A${lastRow}`); // Months column
    const salesRange = sheet.getRange(`C1:C${lastRow}`); // Sales column
    const totalCostRange = sheet.getRange(`J1:J${lastRow}`); // Total Cost column
    const profitRange = sheet.getRange(`K1:K${lastRow}`); // Profit column
    
    // Create a new chart
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN) // Set chart type to BAR
      .addRange(range)
      .addRange(salesRange)
      .addRange(totalCostRange)
      .addRange(profitRange)
      .setPosition(1, 13, 0, 0) // Set chart position (row, column, offsetX, offsetY)
      .setOption('title', 'Sales, Total Cost, and Profit by Month') // Set chart title
      .setOption('hAxis.title', 'Values') // Set horizontal axis title
      .setOption('vAxis.title', 'Month') // Set vertical axis title
      .setOption('series', {
        0: { color: '#1f77b4', label: 'Sales' }, // Sales color
        1: { color: '#ff7f0e', label: 'Total Cost' }, // Total Cost color
        2: { color: '#2ca02c', label: 'Profit' }  // Profit color
      })
      .setOption('legend.position', 'bottom') // Position legend at the bottom
      .setOption('isStacked', false) // Not stacked, showing side-by-side bars
      .build();
    
    // Insert the chart into the sheet
    sheet.insertChart(chart);
  }
  
  function generatePieChart() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();
    
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt('Enter the month (YYYY-MM) to generate the pie chart for:', ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() == ui.Button.OK) {
      var month = response.getResponseText();
      var purchases = 0, rental = 0, transportation = 0, others = 0;
      
      for (var i = 1; i < data.length; i++) {
        var date = new Date(data[i][0]);
        var dateString = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        
        if (dateString.startsWith(month)) {
          Logger.log('Row ' + i + ': ' + data[i]);
          
          purchases += parseValue(data[i][2]);
          rental += parseValue(data[i][3]);
          transportation += parseValue(data[i][4]);
          others += parseValue(data[i][5]) + parseValue(data[i][6]);
        }
      }
      
      Logger.log('Purchases: ' + purchases);
      Logger.log('Rental: ' + rental);
      Logger.log('Transportation: ' + transportation);
      Logger.log('Others: ' + others);
      
      var chartData = [
        ['Category', 'Amount'],
        ['Purchases', purchases],
        ['Rental', rental],
        ['Transportation', transportation],
        ['Others', others]
      ];
      
      var chartRange = sheet.getRange(1, 8, chartData.length, chartData[0].length);
      chartRange.setValues(chartData);
      
      var chart = sheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(chartRange)
        .setOption('title', 'Total Cost')
        .setOption('pieSliceText', 'value')
        .setPosition(1, 8, 0, 0)
        .build();
      
      sheet.insertChart(chart);
    }
  }
  
  function parseValue(value) {
    if (typeof value === 'string' && value) {
      return parseFloat(value.replace('$', '').replace(',', ''));
    } else if (typeof value === 'number') {
      return value;
    } else {
      return 0;
    }
  }
  
  // Create line chart
  function generateMultiLineChart() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();
    
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt('Enter the start month and year (MM-YYYY) and the two headers to plot (comma-separated):', ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() == ui.Button.OK) {
      var input = response.getResponseText().split(',');
      if (input.length < 3) {
        ui.alert('Please enter the start month/year and two headers, separated by commas.');
        return;
      }
      
      var startMonthYear = input[0].trim();
      var header1 = input[1].trim();
      var header2 = input[2].trim();
      
      var [startMonth, startYear] = startMonthYear.split('-').map(Number);
      
      var header1Index = data[0].indexOf(header1);
      var header2Index = data[0].indexOf(header2);
      
      if (header1Index === -1 || header2Index === -1) {
        ui.alert('Invalid headers. Please make sure the headers are correct.');
        return;
      }
      
      var header1Data = [];
      var header2Data = [];
      var months = [];
      
      for (var i = 1; i < data.length; i++) {
        var month = data[i][0];
        var year = data[i][1];
        
        if ((year === startYear && month >= startMonth) || (year === startYear + 1 && month < startMonth)) {
          header1Data.push(data[i][header1Index]);
          header2Data.push(data[i][header2Index]);
          months.push(month);
        }
        
        if (months.length === 12) break;
      }
      
      var chartData = [['Month', header1, header2]];
      for (var j = 0; j < months.length; j++) {
        chartData.push([months[j], header1Data[j], header2Data[j]]);
      }
      
      var chartRange = sheet.getRange(1, 15, chartData.length, chartData[0].length);
      chartRange.setValues(chartData);
      
      var chart = sheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(chartRange)
        .setOption('title', header1 + ' vs ' + header2)
        .setOption('hAxis', {title: 'Month'})
        .setOption('vAxis', {title: 'Amount'})
        .setOption('series', {
          0: {color: 'blue'},
          1: {color: 'red'}
        })
        .setPosition(5, 5, 0, 0)
        .build();
      
      sheet.insertChart(chart);
    }
  }
  
  function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Menu')
        .addItem('Enter Data', 'showInputForm')
        .addItem('Calculate Statistic', 'showMonthYearForm')
        .addItem('predict','calculatePredictions')
        .addItem('Generate Column Chart', 'createBarChartForMonthlyData')
        .addItem('Generate Pie Chart', 'generatePieChart')
        .addItem('Generate Multi-Line Chart', 'generateMultiLineChart')
        .addToUi();
  }
  