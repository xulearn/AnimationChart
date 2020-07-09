/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Internal used const. DO NOT CHANGE
// const chartName = "DynamicChart";

let activeTableId;

//for original table
let totalColumnCount;
let totalRowCount;

//for line chart
let lineChartName = "LineChartName";
let linePointSetLabel;
let linePointUnsetLabel;
// let seriesUpdate;

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("createLineChart").onclick = CreateLineChart;
    document.getElementById("playLineChart").onclick = PlayLineChart;

  }
});

export async function CreateLineChart() {
  try {
    await Excel.run(async context => {

      // Find selected table
      const activeRange = context.workbook.getSelectedRange();
      let dataTables = activeRange.getTables(false);
      dataTables.load("items");
      await context.sync();

      // Get active table
      let dataTable = dataTables.items[0];
      let dataSheet = context.workbook.worksheets.getActiveWorksheet();
      activeTableId = dataTable.id;
      let table = dataSheet.tables.getItem(activeTableId);
      await context.sync();

      let wholeRange = table.getRange();
      wholeRange.load("rowCount");
      wholeRange.load("columnCount");
      await context.sync();

      totalColumnCount = wholeRange.columnCount;
      totalRowCount = wholeRange.rowCount;

      //get initial range, at least one category column and two data columns
      let initialCell = table.getRange().getCell(0, 0);
      let initialRange = initialCell.getAbsoluteResizedRange(totalRowCount, 3); //不是偏移，而是写要的绝对量
      //if the category sorted by columns
      let initalCategoryName = table.getRange().getCell(0, 1).getAbsoluteResizedRange(1, 2);

      let lineChart = dataSheet.charts.add(Excel.ChartType.line, initialRange, "Rows");
      lineChart.set({
        name : lineChartName
      });
      let categoryAxis = lineChart.axes.getItem(Excel.ChartAxisType.category);
      categoryAxis.setCategoryNames(initalCategoryName);

      //set serie label
      // seriesUpdate = {
      //   smooth : true
      // };
      let setLabel = {
        showSeriesName: true,
        showValue: true,
        numberFormat: "#,##0"
      };
      let unsetLabel = {
        showSeriesName: false,
        showValue: false
      };
      linePointSetLabel = {
        dataLabel: setLabel
      };
      linePointUnsetLabel = {
        dataLabel: unsetLabel
      };
      for(let i = totalRowCount - 2; i>=0;--i){
        //getItem(1):初始两列，只显示最右一列; totalRowCount - 2:series从0计数，totalRowCount从1技术
        lineChart.series.getItemAt(i).points.getItemAt(1).set(linePointSetLabel);
      }
      await context.sync();

    });
  } catch (error) {
    console.error(error);
  }
}


export async function PlayLineChart() {
  try {
    await Excel.run(async context => {

      let dataSheet = context.workbook.worksheets.getActiveWorksheet();
      let table = dataSheet.tables.getItem(activeTableId);
      let lineChart = dataSheet.charts.getItem(lineChartName);
      let categoryAxis = lineChart.axes.getItem(Excel.ChartAxisType.category);
      await context.sync();

      //get initial range, at least one category column and two data columns
      let initialCell = table.getRange().getCell(0, 0);
      let initialRange = initialCell.getAbsoluteResizedRange(totalRowCount, 3); //不是偏移，而是写要的绝对量
      let initalCategoryName = table.getRange().getCell(0, 1).getAbsoluteResizedRange(1, 2);

      for (let i = 4; i < totalColumnCount; ++i) {  //from the forth column

        let resizedRange = initialRange.getAbsoluteResizedRange(totalRowCount, i); //向右平移一个列单位

        sleep(1000);
        lineChart.setData(resizedRange, "Rows"); //dynamic change chart
  
        //set categoryName
        let resizedNameRange = initalCategoryName.getAbsoluteResizedRange(1, i);
        categoryAxis.setCategoryNames(resizedNameRange);
        
        //remove original datalabel
        for (let j = totalRowCount-2; j >=0; --j) {
          lineChart.series.getItemAt(j).points.getItemAt(i - 3).set(linePointUnsetLabel); //getItemAt(i - 3): pionts从0开始计数
        }
        await context.sync();

        //set new datalabel
        for (let j = totalRowCount-2; j >=0; --j) {
          lineChart.series.getItemAt(j).points.getItemAt(i - 2).set(linePointSetLabel);
        }
        await context.sync();
      }
      await context.sync();
    });
  }catch (error) {
    console.error(error);
  }
}


export async function CreateBarChart() {
  try {
    await Excel.run(async context => {

    });
  }catch (error) {
    console.error(error);
  }
}


function sleep(sleepTime) {
  var start = new Date().getTime();
  for (var i = 0; i < 1e7; i++) {
    if (new Date().getTime() - start > sleepTime) {
      break;
    }
  }
}