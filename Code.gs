
// main function, this is where the code flows
function updateAverages() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //fetch the 24 hour price data 
  const newData = getNew24HourData(ss);

  //append the newData to 'historical data' sheet
  appendHistoricalData(ss, newData);
  
  //get the last 7 and 28 days of price data
  const weeklyData = getLast7DaysPrices(ss);
  const fourWeekData =  getLast4WeekPrices(ss);
  
  //calculate the average prices of both periods
  const weeklyAverages = calculateAverage(weeklyData, 7);
  const fourWeekAverages = calculateAverage(fourWeekData, 28);
  
  // update the 'Averages' sheet with the new weekly and four weeks averages
  updateAveragesSheet(ss, weeklyAverages, fourWeekAverages);

  //generate charts based on the stored data
  createCharts(ss, weeklyAverages, fourWeekAverages);

  Logger.log("Averages have been successfully updated!");
}

//helper function to reduce redundancy 
function getOrCreateSheet(ss, sheetName) {
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
}

//this is where the charts are generated
function createCharts(ss, weeklyAverages, fourWeeksAverage) {
  // fetch both the chart sheets
  const chartSheet = getOrCreateSheet(ss, "Weekly Chart");
  const fourWeeksSheet = getOrCreateSheet(ss, "FourWeeks Chart");

  //make sure the chartsheets exists
  if (!chartSheet || !fourWeeksSheet) {
    Logger.log("Weekly Chart or FourWeeks Chart sheet not found!");
    return;
  }

  const products = Object.keys(weeklyAverages); //extract products
  const headers = ["Date", ...products]; // defining the chart header

  // if the chart sheets are empty, set headers
  if (chartSheet.getLastRow() === 0) {
    chartSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  if (fourWeeksSheet.getLastRow() === 0) {
    fourWeeksSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }


  const now = new Date();
  const formattedDate = Utilities.formatDate(now, "America/Toronto", "yyyy/MM/dd"); //Change TimeZone Here



  // Updates or appends data to a given sheet. If the data already exists it updates; otherswise, it appends.

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
    
    // Prevent duplicate data entry for the same date
    if (lastDate === formattedDate) {
      sheet.getRange(lastRow, 1, 1, plainTextData.length).setValues([plainTextData]);
      Logger.log(`Updated today's existing data in ${sheet.getName()}.`);
    } else {
      sheet.appendRow(plainTextData);
      Logger.log(`New data added to ${sheet.getName()}.`);
    }
  }

  // prepare and store data for both weekly and four-week charts
  const weeklyRowData = [
    formattedDate,
    ...products.map(p => parseFloat(weeklyAverages[p]).toFixed(2).toString()),
  ];
  const fourWeeksRowData = [
    formattedDate,
    ...products.map(p => parseFloat(fourWeeksAverage[p]).toFixed(2).toString()),
  ];

  // add data to the chart sheets
  updateOrAppend(chartSheet, weeklyRowData);
  updateOrAppend(fourWeeksSheet, fourWeeksRowData);

  //get the range for both sheet
  const weeklyDataRange = chartSheet.getRange(2, 1, chartSheet.getLastRow() - 1, headers.length);
  
  const fourWeeksDataRange = fourWeeksSheet.getRange(2, 1, fourWeeksSheet.getLastRow() - 1, headers.length);


  // get the date for both the sheets data to use as x-axis labels on the chart
  const weeklyDates = chartSheet.getRange(2, 1, chartSheet.getLastRow() - 1, 1).getValues().map(row => row[0].toString());
  const fourWeeksDates = fourWeeksSheet.getRange(2, 1, fourWeeksSheet.getLastRow() - 1, 1).getValues().map(row => row[0].toString());

  //remove any existing charts
  chartSheet.getCharts().forEach(chart => chartSheet.removeChart(chart));
  fourWeeksSheet.getCharts().forEach(chart => fourWeeksSheet.removeChart(chart));

  // craete an object to hold the series option for both the weekly and 4 weeks chart chart
  const weeklySeriesOptions = {};
  products.forEach((product, index) => {
    weeklySeriesOptions[index] = { labelInLegend: product };
  });

  const fourWeeksSeriesOptions = {}; 
  products.forEach((product, index) => {
    fourWeeksSeriesOptions[index] = { labelInLegend: product };
  });

  
  //create the charts
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

// get the 24 hour prices 
function getNew24HourData(ss) {
  // get or create the sheet "New 24 Hour Data"
  const sheet = getOrCreateSheet(ss, "New 24 Hour Data");
  
  // check if the sheet exist if not, then return empty array
  if (!sheet) {
    Logger.log("Sheet not found!");
    return [];
  }
  
  // get the headers from the first row of the sheet, the column titles
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // get the last row
  const lastRow = sheet.getLastRow();

  // find the column index of both column title
  const priceColumnIndex = headers.indexOf("24 Hour Price") + 1;
  const productColumnIndex = headers.indexOf("Product") + 1;
  
  //get the values for both columns
  const valuesProduct = sheet.getRange(2, productColumnIndex, lastRow - 1, 1).getValues();
  const valuesPrice = sheet.getRange(2, priceColumnIndex, lastRow - 1, 1).getValues();
  
  // combine both data into single array of objects
  const combinedData = valuesProduct.map((product, index) => ({
    product: product[0],
    price: valuesPrice[index][0],
  }));
  
  return combinedData;
}

// calculation of averages
function calculateAverage(lastData, requiredDays) {  
  // create an object to hold the averages for both weekly and four weeks
  const averages = {};

  // iterate through each product
  for (const [product, prices] of Object.entries(lastData)) {
    // Check if there are enough prices available for the current product
    if (!prices || prices.length < requiredDays) {
      Logger.log(`Not enough data for product: ${product}. Found ${prices ? prices.length : 0} days, required ${requiredDays}`);
      averages[product] = null;
      continue;
    }

    // calculate the total of the prices by summing up all the values
    const total = prices.reduce((sum, price) => sum + parseFloat(price || 0), 0);
    averages[product] = total / prices.length;
  }

  return averages;
}

// get the recent 4 week prices
function getLast4WeekPrices(ss) {
  // retrieve the 'Historical Data' sheet
  const sheet = getOrCreateSheet(ss, "Historical Data");

  //get all the data from the sheet
  const data = sheet.getDataRange().getValues();

  //destructure the data into headerRow and remaining data
  const [headerRow, ...rows] = data;
  
  const productIndex = headerRow.indexOf("Product");
  const dateIndex = headerRow.indexOf("Date");
  const priceIndex = headerRow.indexOf("24 Hour Prices");

  const fourWeeksAgo = new Date();
  fourWeeksAgo.setDate(fourWeeksAgo.getDate() - (7*4));

  const lastFourWeeksData = {};

  // iteratate throuch each row of data
  rows.forEach(row => {
    const product = row[productIndex]; // product name
    const date = new Date(row[dateIndex]); // converted date
    const price = row[priceIndex]; // price

    // if the date is within the last 4 weeks, add the price to its corresponding product
    if (date >= fourWeeksAgo) {
      if (!lastFourWeeksData[product]) {
        lastFourWeeksData[product] = [];
      }
      lastFourWeeksData[product].push(price);
    }
  });

  // Check each product to ensure it has at least 28 records (4 weeks worth of data)
  for (const product in lastFourWeeksData) {
    if (lastFourWeeksData[product].length < 28) {
      Logger.log(`Warning: Not enough data for product "${product}" in the last 7 days. Only ${lastFourWeeksData[product].length}`);
    }
  }

  return lastFourWeeksData;
}

// get the recend 7 days prices
// same functionality with getLast4Weeks function but 7 days
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

//append the retrieved 24 hour price to historical data sheet
function appendHistoricalData(ss, combinedData) {
  // retrieve the "Historical Data" sheet or create it if it doesn't exist
  const sheet = getOrCreateSheet(ss, "Historical Data");
  if (!sheet) {
    Logger.log("Sheet not found!");
    return;
  }

  // get the current date and time to append to the sheet
  const now = new Date();
  // adjust timezone here
  const formattedDate = Utilities.formatDate(now, "America/Toronto", "yyyy/MM/dd");
  const formattedTime = Utilities.formatDate(now, "America/Toronto", "HH:mm");
  
  // retrieve the headers and the last row of the sheet
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = sheet.getLastRow();
  const dateColumnIndex = headers.indexOf("Date") + 1; 

  // check if data for the current date already exists in the sheet 
  if (lastRow > 1) { 
    const lastRowDate = sheet.getRange(lastRow, dateColumnIndex).getValue(); 
    const normalizedLastRowDate = Utilities.formatDate(
          new Date(lastRowDate), 
          "Asia/Manila", 
          "yyyy/MM/dd"
        );      // If data for today already exists, log the message and stop the function
    if (normalizedLastRowDate === formattedDate) {
      Logger.log(`Data for ${formattedDate} already exists in the last row.`);
      return;
    }

    Logger.log(`Formatted Date: ${formattedDate}, normalizeLastRowDate: ${normalizedLastRowDate}`)
  }

  // append combined data to historical data sheet
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

// update the averages sheet
function updateAveragesSheet(ss, weeklyAverages, fourWeekAverages) {
  // retrieve the "Averages" sheet or create it if it doesn't exist
  const averagesSheet = getOrCreateSheet(ss, "Averages");
  if (!averagesSheet) {
    Logger.log("Averages sheet not found!");
    return;
  }
  
  const lastRow = averagesSheet.getLastRow();

  // If there are any existing data rows, clear them
  if (lastRow > 1) {
    averagesSheet.getRange(2, 1, lastRow - 1, averagesSheet.getLastColumn()).clear();
  }
  
  const headers = averagesSheet.getRange(1, 1, 1, averagesSheet.getLastColumn()).getValues()[0];
  const productCol = headers.indexOf("Product") + 1;
  const weeklyCol = headers.indexOf("Weekly Average Price") + 1;
  const fourWeekCol = headers.indexOf("4 Week Average Price") + 1;
  
  let row = 2;

  // append data
  for (const product in weeklyAverages) {
    averagesSheet.getRange(row, productCol).setValue(product);
    
    // Check if weekly average exists (has enough data)
    if (weeklyAverages[product] !== null) {
      averagesSheet.getRange(row, weeklyCol).setValue(weeklyAverages[product].toFixed(2));
    } else {
      // Leave cell empty if not enough data
      averagesSheet.getRange(row, weeklyCol).setValue("");
      Logger.log(`Not enough weekly data for product "${product}", leaving cell empty.`);
    }
    
    // Check if 4-week average exists (has enough data)
    if (fourWeekAverages[product] !== null) {
      averagesSheet.getRange(row, fourWeekCol).setValue(fourWeekAverages[product].toFixed(2));
    } else {
      // Leave cell empty if not enough data
      averagesSheet.getRange(row, fourWeekCol).setValue("");
      Logger.log(`Not enough 4-week data for product "${product}", leaving cell empty.`);
    }
    
    row++;
  }
}
