
// main function, this is where the code flows
function updateAverages() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const newData = getNew24HourData(ss);

  appendHistoricalData(ss, newData);
  
  const weeklyData = getLast7DaysPrices(ss);
  const fourWeekData =  getLast4WeekPrices(ss);
  
  const weeklyAverages = calculateAverage(weeklyData);
  const fourWeekAverages = calculateAverage(fourWeekData);
  
  updateAveragesSheet(ss, weeklyAverages, fourWeekAverages);

  Logger.log(`Weekly ${weeklyAverages}, 4 Weeks ${fourWeekAverages}`);

  createCharts(ss, weeklyAverages, fourWeekAverages);

  Logger.log("Averages have been successfully updated!");
}

//helper function to reduce redundancy 
function getOrCreateSheet(ss, sheetName) {
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
}

//this is where the charts are generated
function createCharts(ss, weeklyAverages, fourWeeksAverage) {
  const chartSheet = getOrCreateSheet(ss, "Weekly Chart");
  const fourWeeksSheet = getOrCreateSheet(ss, "FourWeeks Chart");

  if (!chartSheet || !fourWeeksSheet) {
    Logger.log("Weekly Chart or FourWeeks Chart sheet not found!");
    return;
  }

  const products = Object.keys(weeklyAverages);
  const headers = ["Date", ...products];

  if (chartSheet.getLastRow() === 0) {
    chartSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  if (fourWeeksSheet.getLastRow() === 0) {
    fourWeeksSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const now = new Date();
  const formattedDate = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "yyyy/MM/dd");

  function updateOrAppend(sheet, data) { 
    const lastRow = sheet.getLastRow();
    const lastDate = lastRow > 1
      ? Utilities.formatDate(
          new Date(sheet.getRange(lastRow, 1).getValue()), // Retrieve and convert to Date
          ss.getSpreadsheetTimeZone(),
          "yyyy/MM/dd"
        )
      : "";

    const plainTextData = data.map(value => value.toString());
    // convert into plainText since charts not working well with Date Obj
    if (lastDate === formattedDate) {
      sheet.getRange(lastRow, 1, 1, plainTextData.length).setValues([plainTextData]);
      Logger.log(`Updated today's existing data in ${sheet.getName()}.`);
    } else {
      sheet.appendRow(plainTextData);
      Logger.log(`New data added to ${sheet.getName()}.`);
    }
  }

  const weeklyRowData = [
    formattedDate,
    ...products.map(p => parseFloat(weeklyAverages[p]).toFixed(2).toString()),
  ];
  const fourWeeksRowData = [
    formattedDate,
    ...products.map(p => parseFloat(fourWeeksAverage[p]).toFixed(2).toString()),
  ];

  updateOrAppend(chartSheet, weeklyRowData);
  updateOrAppend(fourWeeksSheet, fourWeeksRowData);

  const weeklyDataRange = chartSheet.getRange(2, 1, chartSheet.getLastRow() - 1, headers.length);
  
  const fourWeeksDataRange = fourWeeksSheet.getRange(2, 1, fourWeeksSheet.getLastRow() - 1, headers.length);


  const weeklyDates = chartSheet.getRange(2, 1, chartSheet.getLastRow() - 1, 1).getValues().map(row => row[0].toString());
  const fourWeeksDates = fourWeeksSheet.getRange(2, 1, fourWeeksSheet.getLastRow() - 1, 1).getValues().map(row => row[0].toString());

  chartSheet.getCharts().forEach(chart => chartSheet.removeChart(chart));
  fourWeeksSheet.getCharts().forEach(chart => fourWeeksSheet.removeChart(chart));

  const weeklySeriesOptions = {};
  products.forEach((product, index) => {
    weeklySeriesOptions[index] = { labelInLegend: product };
  });

  const fourWeeksSeriesOptions = {}; 
  products.forEach((product, index) => {
    fourWeeksSeriesOptions[index] = { labelInLegend: product };
  });

  const weeklyChart = chartSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(weeklyDataRange)
    .setPosition(2, headers.length + 2, 0, 0)
    .setOption("title", "Weekly Average Prices")
    .setOption("useFirstRowAsHeaders", true) 
    .setOption("hAxis", { 
      title: "Date", 
      slantedText: true,
      format: 'yyyy/MM/dd',
      ticks: weeklyDates.slice(1),
      gridlines: { count: -1 },
      minorGridlines: { count: 0 }
    })
    .setOption("vAxis", { title: "Weekly Average Price" })
    .setOption("legend", { position: 'top' })
    .setOption("series", weeklySeriesOptions)
    .setOption("pointSize", 5)
    .setOption("lineWidth", 2)
    .build();

  chartSheet.insertChart(weeklyChart);

  const fourWeeksChart = fourWeeksSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(fourWeeksDataRange)
    .setPosition(2, headers.length + 2, 0, 0)
    .setOption("title", "Four Weeks Average Prices")
    .setOption("useFirstRowAsHeaders", true) 
    .setOption("hAxis", { 
      title: "Date", 
      slantedText: true,
      format: 'yyyy/MM/dd',
      ticks: fourWeeksDates.slice(1), // Start x-axis ticks from the second row
      gridlines: { count: -1 },
      minorGridlines: { count: 0 }
    })
    .setOption("vAxis", { title: "Four Weeks Average Price" })
    .setOption("legend", { position: 'top' })
    .setOption("series", fourWeeksSeriesOptions)
    .setOption("pointSize", 5)
    .setOption("lineWidth", 2)
    .build();

    fourWeeksSheet.insertChart(fourWeeksChart);


    Logger.log("Weekly and Four Weeks Average Charts Created Successfully.");
}


function getNew24HourData(ss) {
  const sheet = getOrCreateSheet(ss, "New 24 Hour Data");
  
  if (!sheet) {
    Logger.log("Sheet not found!");
    return [];
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = sheet.getLastRow();
  const priceColumnIndex = headers.indexOf("24 Hour Price") + 1;
  const productColumnIndex = headers.indexOf("Product") + 1;
  
  const valuesProduct = sheet.getRange(2, productColumnIndex, lastRow - 1, 1).getValues();
  const valuesPrice = sheet.getRange(2, priceColumnIndex, lastRow - 1, 1).getValues();
  
  const combinedData = valuesProduct.map((product, index) => ({
    product: product[0],
    price: valuesPrice[index][0],
  }));
  
  return combinedData;
}

function calculateAverage(lastData) {  
  const weeklyAverages = {};

  for (const [product, prices] of Object.entries(lastData)) {
    if (prices.length === 0) {
      Logger.log(`No prices available for product: ${product}`);
      weeklyAverages[product] = null;
      continue;
    }

    const total = prices.reduce((sum, price) => sum + parseFloat(price || 0), 0);
    weeklyAverages[product] = total / prices.length;
  }

  return weeklyAverages;
}

function getLast4WeekPrices(ss) {
  const sheet = getOrCreateSheet(ss, "Historical Data");
  const data = sheet.getDataRange().getValues();
  const [headerRow, ...rows] = data;
  
  const productIndex = headerRow.indexOf("Product");
  const dateIndex = headerRow.indexOf("Date");
  const priceIndex = headerRow.indexOf("24 Hour Prices");

  const fourWeeksAgo = new Date();
  fourWeeksAgo.setDate(fourWeeksAgo.getDate() - (7*4));

  const lastFourWeeksData = {};

  rows.forEach(row => {
    const product = row[productIndex];
    const date = new Date(row[dateIndex]);
    const price = row[priceIndex];

    if (date >= fourWeeksAgo) {
      if (!lastFourWeeksData[product]) {
        lastFourWeeksData[product] = [];
      }
      lastFourWeeksData[product].push(price);
    }
  });

  for (const product in lastFourWeeksData) {
    if (lastFourWeeksData[product].length < 28) {
      Logger.log(`Warning: Not enough data for product "${product}" in the last 7 days. Only ${lastFourWeeksData[product].length}`);
    }
  }

  return lastFourWeeksData;
}

function getLast7DaysPrices(ss) {
  const sheet = getOrCreateSheet(ss, "Historical Data");
  const data = sheet.getDataRange().getValues();
  const [headerRow, ...rows] = data;
  
  const productIndex = headerRow.indexOf("Product");
  const dateIndex = headerRow.indexOf("Date");
  const priceIndex = headerRow.indexOf("24 Hour Prices");

  const sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);

  const last7DaysData = {};

  rows.forEach(row => {
    const product = row[productIndex];
    const date = new Date(row[dateIndex]);
    const price = row[priceIndex];

    if (date >= sevenDaysAgo) {
      if (!last7DaysData[product]) {
        last7DaysData[product] = [];
      }
      last7DaysData[product].push(price);
    }
  });

  for (const product in last7DaysData) {
    if (last7DaysData[product].length < 7) {
      Logger.log(`Warning: Not enough data for product "${product}" in the last 7 days. Only ${last7DaysData[product].length}`);
    }
  }

  return last7DaysData;
}

function appendHistoricalData(ss, combinedData) {
  const sheet = getOrCreateSheet(ss, "Historical Data");
  if (!sheet) {
    Logger.log("Sheet not found!");
    return;
  }

  const now = new Date();
  const formattedDate = `${now.getFullYear()}/${now.getMonth() + 1}/${now.getDate()}`;
  const formattedTime = `${(now.getHours() < 10 ? "0" : "") + now.getHours()}:${(now.getMinutes() < 10 ? "0" : "") + now.getMinutes()}`;

  

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = sheet.getLastRow();
  const dateColumnIndex = headers.indexOf("Date") + 1; 

  if (lastRow > 1) { 
    const lastRowDate = sheet.getRange(lastRow, dateColumnIndex).getValue(); 
    const normalizedLastRowDate = `${lastRowDate.getFullYear()}/${lastRowDate.getMonth() + 1}/${lastRowDate.getDate()}`;

    if (normalizedLastRowDate === formattedDate) {
      Logger.log(`Data for ${formattedDate} already exists in the last row.`);
      return;
    }
  }

  let currentRow = lastRow + 1;
  combinedData.forEach(item => {
    sheet.getRange(currentRow, headers.indexOf("Date") + 1).setValue(formattedDate);
    sheet.getRange(currentRow, headers.indexOf("Time") + 1).setValue(formattedTime);
    sheet.getRange(currentRow, headers.indexOf("Product") + 1).setValue(item.product);
    sheet.getRange(currentRow, headers.indexOf("24 Hour Prices") + 1).setValue(item.price);
    currentRow++;
  });

  Logger.log(`Data appended successfully for ${formattedDate}.`);
}


function updateAveragesSheet(ss, weeklyAverages, fourWeekAverages) {
  const averagesSheet = getOrCreateSheet(ss, "Averages");
  if (!averagesSheet) {
    Logger.log("Averages sheet not found!");
    return;
  }
  
  const lastRow = averagesSheet.getLastRow();
  if (lastRow > 1) {
    averagesSheet.getRange(2, 1, lastRow - 1, averagesSheet.getLastColumn()).clear();
  }
  
  const headers = averagesSheet.getRange(1, 1, 1, averagesSheet.getLastColumn()).getValues()[0];
  const productCol = headers.indexOf("Product") + 1;
  const weeklyCol = headers.indexOf("Weekly Average Price") + 1;
  const fourWeekCol = headers.indexOf("4 Week Average Price") + 1;
  
  let row = 2;
  for (const product in weeklyAverages) {
    averagesSheet.getRange(row, productCol).setValue(product);
    averagesSheet.getRange(row, weeklyCol).setValue(weeklyAverages[product].toFixed(2));
    averagesSheet.getRange(row, fourWeekCol).setValue(fourWeekAverages[product].toFixed(2));
    row++;
  }
}
