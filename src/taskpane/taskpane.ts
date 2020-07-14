/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Internal used const. DO NOT CHANGE
let activeTableId;
let inputPointItems: any;
let orientation: number;

let pointItemsCount: number;

//for original table
let totalColumnCount: number;
let totalRowCount: number;

//for line chart
let lineChartName = "LineChartName";
let linePointSetLabel;
let linePointUnsetLabel;
// let seriesUpdate;

//--------------------------------
// Parameters. Modify it if needed.
const chartWidth = 750,
  chartHeight = 550,
  chartLeft = 150,
  chartTop = 50;
const splitIncreasement = 2;
const colorList = [
  "#afc97a",
  "#cd7371",
  "#729aca",
  "#b65708",
  "#276a7c",
  "#4d3b62",
  "#5f7530",
  "#772c2a",
  "#2c4d75",
  "#f79646",
  "#4bacc6",
  "#8064a2",
  "#9bbb59",
  "#c0504d",
  "#4f81bd"
];
const fontSize_Title = 28,
  fontSize_CategoryName = 13,
  fontSize_AxisValue = 11,
  fontSize_DataLabel = 13;

// Internal used const. DO NOT CHANGE
//for barChart and columnChart and tool
const toolSheetName = "toolSheet"; //+UUID
const toolTableName = "toolTable"; //+UUID

const barChartName = "BarChartName";
const columnChartName = "ColumnChartName";

const barChartFlag = 1;
const columnChartFlag = 2;

enum ToolTableColumnIndex {
  category = 0,
  value = 1,
  color = 2,
  map = 3
}

enum ToolTableColumnName {
  category = "categoryColumn",
  //value
  color = "colorColumn",
  map = "mapColumn"
}

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    //line chart
    document.getElementById("createLineChart").onclick = CreateLineChart;
    document.getElementById("playLineChart").onclick = PlayLineChart;
    //bar chart
    document.getElementById("createBarChart").onclick = CreateBarChart;
    document.getElementById("playBarChart").onclick = PlayBarChart;
    //column chart
    document.getElementById("createColumnChart").onclick = CreateColumnChart;
    document.getElementById("playColumnChart").onclick = PlayColumnChart;
  }
});

/**
 * create line chart
 */
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
      let initialRange = initialCell.getAbsoluteResizedRange(totalRowCount, 3); //It's not an offset, it's an absolute amount .
      //if the category sorted by columns
      let initalCategoryName = table
        .getRange()
        .getCell(0, 1)
        .getAbsoluteResizedRange(1, 2);

      let lineChart = dataSheet.charts.add(Excel.ChartType.line, initialRange, "Rows");
      // lineChart.set({
      //   name: lineChartName
      // });
      let lineChartHeight = chartHeight - 50;
      lineChart.set({ name: lineChartName, height: lineChartHeight, width: chartWidth, left: chartLeft, top: chartTop });
      let categoryAxis = lineChart.axes.getItem(Excel.ChartAxisType.category);
      categoryAxis.setCategoryNames(initalCategoryName);

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
      for (let i = totalRowCount - 2; i >= 0; --i) {
        //getItem(1):The first two columns, only the rightmost column is displayed.
        //totalRowCount - 2:series counts from 0, totalRowCount counts from 1.
        lineChart.series
          .getItemAt(i)
          .points.getItemAt(1)
          .set(linePointSetLabel);
      }
      lineChart.title.text = "LineChart";
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

/**
 * play line chart
 */
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
      let initialRange = initialCell.getAbsoluteResizedRange(totalRowCount, 3); //It's not an offset, it's an absolute amount .
      let initalCategoryName = table
        .getRange()
        .getCell(0, 1)
        .getAbsoluteResizedRange(1, 2);

      for (let i = 4; i <= totalColumnCount; ++i) { // 4th column
        //todo!!
        //from the forth column

        let resizedRange = initialRange.getAbsoluteResizedRange(totalRowCount, i); //Translate one column unit to the right.

        //todo: do not use sleep
        // sleep(1000);
        lineChart.setData(resizedRange, "Rows"); //dynamic change chart

        //set categoryName
        let resizedNameRange = initalCategoryName.getAbsoluteResizedRange(1, i);
        categoryAxis.setCategoryNames(resizedNameRange);

        //remove original datalabel
        for (let j = totalRowCount - 2; j >= 0; --j) {
          lineChart.series
            .getItemAt(j)
            .points.getItemAt(i - 3)
            .set(linePointUnsetLabel); //getItemAt(i - 3): pionts count from 0.
        }
        await context.sync();

        //set new datalabel
        for (let j = totalRowCount - 2; j >= 0; --j) {
          lineChart.series
            .getItemAt(j)
            .points.getItemAt(i - 2)
            .set(linePointSetLabel);
        }
        await context.sync();
      }
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

/**
 * create for bar chart
 */
export async function CreateBarChart() {
  try {
    await CreateBarOrColumnChart(barChartFlag);
  } catch (error) {
    console.error(error);
  }
}

/**
 * create for column chart
 */
export async function CreateColumnChart() {
  try {
    await CreateBarOrColumnChart(columnChartFlag);
  } catch (error) {
    console.error(error);
  }
}

/**
 * create for bar or column chart
 */
export async function CreateBarOrColumnChart(flag: number) {
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
      activeTableId = dataTable.id; //id can not be loaded
      let table = dataSheet.tables.getItem(activeTableId);
      await context.sync();

      let wholeRange = table.getRange();
      wholeRange.load("rowCount");
      wholeRange.load("columnCount");
      await context.sync();
      totalColumnCount = wholeRange.columnCount;
      totalRowCount = wholeRange.rowCount;

      //create toolTable
      //delete the old chart and sheet
      let toolSheet: Excel.Worksheet;
      toolSheet = context.workbook.worksheets.getItemOrNullObject(toolSheetName);
      toolSheet.load();
      await context.sync();
      let lastBarChart: Excel.Chart;
      let lastColumnChart: Excel.Chart;

      if (JSON.stringify(toolSheet) !== "{}") {
        lastBarChart = dataSheet.charts.getItemOrNullObject(barChartName);
        lastColumnChart = dataSheet.charts.getItemOrNullObject(columnChartName);
        //chart delete
        lastBarChart.load();
        lastColumnChart.load();
        await context.sync();
        if (JSON.stringify(lastBarChart) !== "{}") {
          lastBarChart.delete();
        }
        if (JSON.stringify(lastColumnChart) !== "{}") {
          lastColumnChart.delete();
        }
        toolSheet.delete();
      }
      toolSheet = context.workbook.worksheets.add(toolSheetName);

      let toolRange = toolSheet.getCell(0, 0).getAbsoluteResizedRange(totalRowCount, 4);
      let toolTable = toolSheet.tables.add(toolRange, true);
      toolTable.set({
        name: toolTableName
      });
      //set columnName
      toolTable.columns.getItemAt(ToolTableColumnIndex.category).set({ name: ToolTableColumnName.category });
      toolTable.columns.getItemAt(ToolTableColumnIndex.color).set({ name: ToolTableColumnName.color });
      toolTable.columns.getItemAt(ToolTableColumnIndex.map).set({ name: ToolTableColumnName.map });

      let categoryBodyRange = toolTable.columns.getItem(ToolTableColumnName.category).getDataBodyRange();
      let curIteratedRange = toolTable.columns.getItemAt(ToolTableColumnIndex.value).getRange();
      let curIteratedBodyRange = toolTable.columns.getItemAt(ToolTableColumnIndex.value).getDataBodyRange();
      let colorBodyRange = toolTable.columns.getItem(ToolTableColumnName.color).getDataBodyRange();
      let mapBodyRange = toolTable.columns.getItem(ToolTableColumnName.map).getDataBodyRange();

      //copy Range
      categoryBodyRange.copyFrom(table.columns.getItemAt(0).getDataBodyRange());
      curIteratedRange.copyFrom(table.columns.getItemAt(1).getRange()); //copy headers too

      colorBodyRange.load("values");
      await context.sync();
      let tmpColorArr = [];
      for (let i = 0; i < totalRowCount - 1; ++i) {
        tmpColorArr.push([colorList[i % colorList.length]]);
      }
      colorBodyRange.values = tmpColorArr;

      mapBodyRange.load("values");
      await context.sync();
      let tmpMapArr = [];
      for (let i = 1; i < totalRowCount; ++i) {
        tmpMapArr.push([i]);
      }
      mapBodyRange.values = tmpMapArr;

      //input
      let inputElement = document.getElementById("PointItems") as HTMLInputElement;
      inputPointItems = inputElement.value;
      let optionElement = document.getElementById("orientation") as HTMLOptionElement;
      orientation = Number(optionElement.value);

      //get input and target items
      pointItemsCount = formatInput(inputPointItems, totalRowCount);

      let targetIteratedBodyRange = getPartialRange(curIteratedBodyRange, pointItemsCount, orientation);
      let targetCategoryRange = getPartialRange(categoryBodyRange, pointItemsCount, orientation);

      // Create Chart
      toolTable.sort.apply([{ key: 1, ascending: true }], true); //toolTable only does ascending sort
      let chart: Excel.Chart;
      if (flag === barChartFlag) {
        chart = dataSheet.charts.add(Excel.ChartType.barClustered, targetIteratedBodyRange);
        chart.set({ name: barChartName, height: chartHeight, width: chartWidth, left: chartLeft, top: chartTop });
      } else {
        chart = dataSheet.charts.add(Excel.ChartType.columnClustered, targetIteratedBodyRange);
        chart.set({ name: columnChartName, height: chartHeight, width: chartWidth, left: chartLeft, top: chartTop });
      }

      let curheaderRange = curIteratedRange.getCell(0, 0);
      curheaderRange.load("text");
      await context.sync();
      // Set chart tile and style
      chart.title.text = curheaderRange.text[0][0];
      chart.title.format.font.set({ size: fontSize_Title });
      chart.legend.set({ visible: false });

      // Set Axis
      let categoryAxis = chart.axes.getItem(Excel.ChartAxisType.category);
      categoryAxis.setCategoryNames(targetCategoryRange);
      categoryAxis.set({ visible: true });
      categoryAxis.format.font.set({ size: fontSize_CategoryName });
      let valueAxis = chart.axes.getItem(Excel.ChartAxisType.value);
      valueAxis.format.font.set({ size: fontSize_AxisValue });

      let series = chart.series.getItemAt(0);
      series.set({ hasDataLabels: true, gapWidth: 30 });
      series.dataLabels.set({ showCategoryName: false, numberFormat: "#,##0" });
      series.dataLabels.format.font.set({ size: fontSize_DataLabel });
      series.points.load();
      await context.sync();

      colorBodyRange.load("values");
      await context.sync();
      let sortedColorArr = colorBodyRange.values;

      // Set data points color
      for (let i = 0; i < series.points.count; i++) {
        if (orientation === 1) {
          series.points
            .getItemAt(i)
            .format.fill.setSolidColor(sortedColorArr[totalRowCount - pointItemsCount - 1 + i][0]);
        } else {
          series.points.getItemAt(i).format.fill.setSolidColor(sortedColorArr[i][0]);
        }
      }
      series.points.load();

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

/**
 * play for bar chart
 */
export async function PlayBarChart() {
  try {
    await PlayBarOrColumnChart(barChartFlag);
  } catch (error) {
    console.error(error);
  }
}

/**
 * play for column chart
 */
export async function PlayColumnChart() {
  try {
    await PlayBarOrColumnChart(columnChartFlag);
  } catch (error) {
    console.error(error);
  }
}

/**
 * play bar or column chart
 */
export async function PlayBarOrColumnChart(flag: number) {
  try {
    await Excel.run(async context => {
      let dataSheet = context.workbook.worksheets.getActiveWorksheet();
      let table = dataSheet.tables.getItem(activeTableId);

      //get toolTable
      let toolSheet = context.workbook.worksheets.getItem(toolSheetName);
      // let toolTable = dataSheet.tables.getItem(toolTableName);
      let toolTable = toolSheet.tables.getItem(toolTableName);

      let categoryBodyRange = toolTable.columns.getItemAt(ToolTableColumnIndex.category).getDataBodyRange();
      let curIteratedHeaderRange = toolTable.columns.getItemAt(ToolTableColumnIndex.value).getHeaderRowRange();
      let curIteratedBodyRange = toolTable.columns.getItemAt(ToolTableColumnIndex.value).getDataBodyRange();
      let mapBodyRange = toolTable.columns.getItem(ToolTableColumnName.map).getDataBodyRange();
      let colorBodyRange = toolTable.columns.getItem(ToolTableColumnName.color).getDataBodyRange();

      let chart: Excel.Chart;
      if (flag == barChartFlag) {
        chart = dataSheet.charts.getItem(barChartName);
      } else {
        chart = dataSheet.charts.getItem(columnChartName);
      }
      //todo splitIncreasement input

      categoryBodyRange.load("values");
      curIteratedBodyRange.load("values");
      mapBodyRange.load("values");
      colorBodyRange.load("values");
      await context.sync();

      //initial countryArr
      let countryArray: Country[] = [];
      for (let i = 0; i < totalRowCount - 1; ++i) {
        let curCategory = categoryBodyRange.values[i][0];
        let curValue = curIteratedBodyRange.values[i][0];
        let curMap = mapBodyRange.values[i][0];
        let curColor = colorBodyRange.values[i][0];

        let curCountry = new Country(curCategory, curValue, curMap, 0, curColor);
        countryArray.push(curCountry);
      }

      // paly
      for (let i = 2; i < totalColumnCount; ++i) {
        let nextIteratedHeaderRange = table.columns.getItemAt(i).getHeaderRowRange(); //from table
        let nextIteratedRange = table.columns.getItemAt(i).getRange();

        nextIteratedRange.load("values");
        curIteratedBodyRange.load("values");
        curIteratedHeaderRange.load("text");
        mapBodyRange.load("values");
        await context.sync();

        let nextArr = mapTargetRangeValue(mapBodyRange, nextIteratedRange);
        // Calculate increase based on current value and next value
        let increaseData = calculateIncrease(curIteratedBodyRange.values, nextArr, splitIncreasement);

        for (let j = 0; j < totalRowCount - 1; ++j) {
          countryArray[j].setIncreasement(increaseData[j]);
        }

        for (let step = 1; step <= splitIncreasement; step++) {
          if (step === splitIncreasement) {
            mapBodyRange.load("values");
            await context.sync();
            //The mapRange here is the one that was ordered in the previous 'else', and you'll have to take it again because countryArr already sorted.
            nextArr = mapTargetRangeValue(mapBodyRange, nextIteratedRange);

            for (let j = 0; j < totalRowCount - 1; ++j) {
              countryArray[j].setValue(nextArr[j][0]);
            }
            //set title
            curIteratedHeaderRange.copyFrom(nextIteratedHeaderRange);
          } else {
            // Add increase amount
            for (let j = 0; j < totalRowCount - 1; ++j) {
              countryArray[j].updateIncrease();
            }
          }

          //sort
          countryArray.sort((a: Country, b: Country) => a.value - b.value); //countryArray only does ascending sort

          //set some value to excel Range
          let categoryArray = [];
          let valueArray = [];
          let mapArray = [];
          let colorArray = [];
          for (let j = 0; j < totalRowCount - 1; ++j) {
            categoryArray.push([countryArray[j].name]);
            valueArray.push([countryArray[j].value]); //the chart will use this column
            mapArray.push([countryArray[j].mapColumn]); //this column will be used to map row's number
            colorArray.push([countryArray[j].color]);
          }
          categoryBodyRange.values = categoryArray;
          curIteratedBodyRange.values = valueArray;
          mapBodyRange.values = mapArray;
          colorBodyRange.values = colorArray;
          await context.sync();

          // Set data points color
          let series = chart.series.getItemAt(0);
          series.load("points");
          colorBodyRange.load("values");
          await context.sync();
          let tmpColorArr = colorBodyRange.values;
          for (let k = 0; k < series.points.count; k++) {
            if (orientation === 1) {
              series.points
                .getItemAt(k)
                .format.fill.setSolidColor(tmpColorArr[totalRowCount - pointItemsCount - 1 + k][0]);
            } else {
              series.points.getItemAt(k).format.fill.setSolidColor(tmpColorArr[k][0]);
            }
          }
          series.points.load();
          await context.sync();
        }

        curIteratedHeaderRange.load("text");
        await context.sync();
        chart.title.text = curIteratedHeaderRange.text[0][0];

        await context.sync();
      }

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

function sleep(sleepTime: number) {
  var start = new Date().getTime();
  for (var i = 0; i < 1e7; i++) {
    if (new Date().getTime() - start > sleepTime) {
      break;
    }
  }
}

function formatInput(input: string, rowCount: number): number {
  let pointItemsCount = Number(input);
  if (isNaN(pointItemsCount) || pointItemsCount <= 0 || pointItemsCount > rowCount || String(input).indexOf(".") >= 0) {
    console.log("please input a integer");
    pointItemsCount = rowCount - 1;
  }
  return pointItemsCount;
}

/**
 * @param originalRange : BodyRange
 * @param pointItemsCount : itemscount that u want
 */
function getPartialRange(originalRange: Excel.Range, pointItemsCount: number, orientation: number): Excel.Range {
  let partialRange: Excel.Range;
  if (orientation === 1) {
    //for top n
    partialRange = originalRange
      .getCell(totalRowCount - pointItemsCount - 1, 0)
      .getAbsoluteResizedRange(pointItemsCount, 1);
  } else {
    partialRange = originalRange.getCell(0, 0).getAbsoluteResizedRange(pointItemsCount, 1);
  }

  return partialRange;
}

// To calculate the increase for each step between next data list and current data list
//function calculateIncrease(current: Array<Array<number>>, next: Array<Array<number>>, steps: number) {
function calculateIncrease(current: any[][], next: any[][], steps: number): any[] {
  if (current.length != next.length) {
    console.error("Error! current data length:" + current.length + ", next data length" + next.length + ".");
  }

  let result = [];
  for (let i = 0; i < current.length; i++) {
    let increasement = (next[i][0] - current[i][0]) / steps;
    result[i] = increasement;
  }

  return result;
}

function mapTargetRangeValue(mapRange: Excel.Range, targetRange: Excel.Range): any[][] {
  let targetArr = [];
  let mapArr = mapRange.values;
  for (let j = 0; j < mapArr.length; ++j) {
    let mapIndex = mapArr[j][0];
    let mapVal = targetRange.values[mapIndex][0];
    targetArr.push([mapVal]);
  }
  return targetArr;
}

function hiddenSheet(sheet: Excel.Worksheet) {
  sheet.set({ visibility: "Hidden" });
  // sheet.set({ visibility: "Visible"});
}

class Country {
  name: string;
  value: number;
  mapColumn: number;
  increasement: number;
  color: string;

  constructor(name: string, value: number, mapColumn: number, increasement: number, color: string) {
    this.name = name;
    this.value = value;
    this.mapColumn = mapColumn;
    this.increasement = increasement;
    this.color = color;
  }

  setValue(value: number): void {
    this.value = value;
  }

  setIncreasement(increasement: number) {
    this.increasement = increasement;
  }

  setColor(color: string) {
    this.color = color;
  }

  updateIncrease(): void {
    this.value = this.value + this.increasement;
  }
}
