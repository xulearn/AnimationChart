/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Internal used const. DO NOT CHANGE
// const chartName = "DynamicChart";

let activeTableId;
let inputPointItems: any;
let orientation: number;

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
const barChartName = "BarChartName";

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
      let initialRange = initialCell.getAbsoluteResizedRange(totalRowCount, 3); //不是偏移，而是写要的绝对量
      //if the category sorted by columns
      let initalCategoryName = table
        .getRange()
        .getCell(0, 1)
        .getAbsoluteResizedRange(1, 2);

      let lineChart = dataSheet.charts.add(Excel.ChartType.line, initialRange, "Rows");
      lineChart.set({
        name: lineChartName
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
      for (let i = totalRowCount - 2; i >= 0; --i) {
        //getItem(1):初始两列，只显示最右一列; totalRowCount - 2:series从0计数，totalRowCount从1技术
        lineChart.series
          .getItemAt(i)
          .points.getItemAt(1)
          .set(linePointSetLabel);
      }
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
      let initialRange = initialCell.getAbsoluteResizedRange(totalRowCount, 3); //不是偏移，而是写要的绝对量
      let initalCategoryName = table
        .getRange()
        .getCell(0, 1)
        .getAbsoluteResizedRange(1, 2);

      for (let i = 4; i < totalColumnCount; ++i) {
        //from the forth column

        let resizedRange = initialRange.getAbsoluteResizedRange(totalRowCount, i); //向右平移一个列单位

        //todo: do not use sleep
        sleep(1000);
        lineChart.setData(resizedRange, "Rows"); //dynamic change chart

        //set categoryName
        let resizedNameRange = initalCategoryName.getAbsoluteResizedRange(1, i);
        categoryAxis.setCategoryNames(resizedNameRange);

        //remove original datalabel
        for (let j = totalRowCount - 2; j >= 0; --j) {
          lineChart.series
            .getItemAt(j)
            .points.getItemAt(i - 3)
            .set(linePointUnsetLabel); //getItemAt(i - 3): pionts从0开始计数
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
    await Excel.run(async context => {
      // Find selected table
      const activeRange = context.workbook.getSelectedRange();
      let dataTables = activeRange.getTables(false);
      dataTables.load("items");
      await context.sync();

      // Get active table
      let dataTable = dataTables.items[0];
      let dataSheet = context.workbook.worksheets.getActiveWorksheet();
      activeTableId = dataTable.id; //id不能load！！！
      let table = dataSheet.tables.getItem(activeTableId);
      await context.sync();

      let wholeRange = table.getRange();
      wholeRange.load("rowCount");
      wholeRange.load("columnCount");
      await context.sync();
      totalColumnCount = wholeRange.columnCount;
      totalRowCount = wholeRange.rowCount;

      //create toolTable
      let toolRange = dataSheet.getCell(0, totalColumnCount + 1).getAbsoluteResizedRange(totalRowCount, 4);
      let toolTable = dataSheet.tables.add(toolRange, true);
      toolTable.set({
        name: "toolTable"
      });
      //set columnName
      toolTable.columns.getItemAt(0).set({ name: "categoryColumn" });
      // toolTable.columns.getItemAt(1).set({ name: "curIteratedColumn" });
      toolTable.columns.getItemAt(2).set({ name: "colorColumn" });
      toolTable.columns.getItemAt(3).set({ name: "mapColumn" });
      // toolTable.columns.getItemAt(4).set({name: "increaseColumn"});  //increase不用了

      let categoryBodyRange = toolTable.columns.getItem("categoryColumn").getDataBodyRange();
      let curIteratedRange = toolTable.columns.getItemAt(1).getRange();
      let curIteratedBodyRange = toolTable.columns.getItemAt(1).getDataBodyRange();
      let colorRange = toolTable.columns.getItem("colorColumn").getRange();
      let mapRange = toolTable.columns.getItem("mapColumn").getRange();

      //copy Range
      categoryBodyRange.copyFrom(table.columns.getItemAt(0).getDataBodyRange());
      curIteratedRange.copyFrom(table.columns.getItemAt(1).getRange()); //copy headers too
      colorRange.load("values");
      await context.sync();
      for (let i = 1; i < totalRowCount; ++i) {
        //填补colorRange，
        colorRange.getCell(i, 0).values = [[colorList[(i - 1) % colorList.length]]]; //只通过cell赋值,且一次sync之后就要赋值
      }
      mapRange.load("values");
      await context.sync();
      for (let i = 1; i < totalRowCount; ++i) {
        mapRange.getCell(i, 0).values = [[i]];
      }

      //input
      let inputElement = document.getElementById("PointItems") as HTMLInputElement;
      inputPointItems = inputElement.value;
      let optionElement = document.getElementById("orientation") as HTMLOptionElement;
      orientation = Number(optionElement.value);
      console.log(orientation);

      //get input and target items
      let pointItemsCount = formatInput(inputPointItems, totalRowCount);

      let targetIteratedBodyRange = getPartialRange(curIteratedBodyRange,pointItemsCount,orientation); 
      let targetCategoryRange = getPartialRange(categoryBodyRange,pointItemsCount,orientation); 

      // Create Chart
      // if(orientation === 1){  //for top n
        toolTable.sort.apply([{ key: 1, ascending: true }], true);  //toolTable only does ascending sort
      // }else{  //for bottom n
      //   toolTable.sort.apply([{ key: 1, ascending: true }], true);
      // }
      // let barChart = dataSheet.charts.add(Excel.ChartType.barClustered, curIteratedBodyRange);
      let barChart = dataSheet.charts.add(Excel.ChartType.barClustered, targetIteratedBodyRange);


      barChart.set({ name: barChartName, height: chartHeight, width: chartWidth, left: chartLeft, top: chartTop });
      let curheaderRange = curIteratedRange.getCell(0, 0);
      curheaderRange.load("text");
      await context.sync();
      // Set chart tile and style
      barChart.title.text = curheaderRange.text[0][0];
      barChart.title.format.font.set({ size: fontSize_Title });
      barChart.legend.set({ visible: false });

      // Set Axis
      let categoryAxis = barChart.axes.getItem(Excel.ChartAxisType.category);
      categoryAxis.setCategoryNames(targetCategoryRange);
      categoryAxis.set({ visible: true });
      categoryAxis.format.font.set({ size: fontSize_CategoryName });
      let valueAxis = barChart.axes.getItem(Excel.ChartAxisType.value);
      valueAxis.format.font.set({ size: fontSize_AxisValue });

      let series = barChart.series.getItemAt(0);
      series.set({ hasDataLabels: true, gapWidth: 30 });
      series.dataLabels.set({ showCategoryName: false, numberFormat: "#,##0" });
      series.dataLabels.format.font.set({ size: fontSize_DataLabel });
      series.points.load();
      await context.sync();

      // Set data points color
      for (let i = 0; i < series.points.count; i++) {
        series.points.getItemAt(i).format.fill.setSolidColor(colorList[i % colorList.length]);
      }
      series.points.load();

      //hide the Range of toolTable
      // toolRange.set({ columnHidden: true });

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

/**
 * play bar chart
 */
export async function PlayBarChart() {
  try {
    await Excel.run(async context => {
      let dataSheet = context.workbook.worksheets.getActiveWorksheet();
      let table = dataSheet.tables.getItem(activeTableId);

      //get toolTable
      let toolTable = dataSheet.tables.getItem("toolTable");
      let categoryBodyRange = toolTable.columns.getItemAt(0).getDataBodyRange();
      let curIteratedHeaderRange = toolTable.columns.getItemAt(1).getHeaderRowRange();
      let curIteratedBodyRange = toolTable.columns.getItemAt(1).getDataBodyRange();
      let mapBodyRange = toolTable.columns.getItem("mapColumn").getDataBodyRange();
      let colorBodyRange = toolTable.columns.getItem("colorColumn").getDataBodyRange();

      let barChart = dataSheet.charts.getItem(barChartName);
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

      //test
      // for(let i=0;i<totalRowCount-1;++i){
      //   console.log(countryArray[i].color);
      // }


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
            //这里的mapRange是上一轮else中排过序的，要再拿一次，因为countryArr已经排序了
            nextArr = mapTargetRangeValue(mapBodyRange, nextIteratedRange);

            for (let j = 0; j < totalRowCount - 1; ++j) {
              countryArray[j].setValue(nextArr[j][0]);
            }

            // //sort
            // countryArray.sort((a: Country, b: Country) => a.value - b.value);
            //set title
            curIteratedHeaderRange.copyFrom(nextIteratedHeaderRange);
          } else {
            // Add increase amount
            for (let j = 0; j < totalRowCount - 1; ++j) {
              countryArray[j].updateIncrease();
            }

            // //sort
            // countryArray.sort((a: Country, b: Country) => a.value - b.value);
          }

          //sort
          // if(orientation === 1){  //for top n
          //   countryArray.sort((a: Country, b: Country) => b.value - a.value);
          // }else{  //for bottom n
            countryArray.sort((a: Country, b: Country) => a.value - b.value); //countryArray only does ascending sort
          // }

          //set some value to excel Range
          let categoryArray = [];
          let valueArray = [];
          let mapArray = [];
          let colorArray = [];
          for (let j = 0; j < totalRowCount - 1; ++j) {
            categoryArray.push([countryArray[j].name]);
            valueArray.push([countryArray[j].value]); //该列必须存，chart用
            mapArray.push([countryArray[j].mapColumn]); //该列用于行间映射
            colorArray.push([countryArray[j].color]);
          }
          categoryBodyRange.values = categoryArray;
          curIteratedBodyRange.values = valueArray;
          mapBodyRange.values = mapArray;
          colorBodyRange.values = colorArray;

          await context.sync();

          // Set data points color
          let series = barChart.series.getItemAt(0);
          series.load("points");
          colorBodyRange.load("values");
          await context.sync();
          for (let k = 0; k < series.points.count; k++) {   //todo：要调整适应top n和bottom n
            series.points.getItemAt(k).format.fill.setSolidColor(colorBodyRange.values[k][0]);
          }

          await context.sync();
        }

        curIteratedHeaderRange.load("text");
        await context.sync();
        barChart.title.text = curIteratedHeaderRange.text[0][0];

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

function formatInput(input: string, rowCount: number): number{
  let pointItemsCount = Number(input);
  if(isNaN(pointItemsCount) || pointItemsCount <= 0 || pointItemsCount > rowCount || String(input).indexOf(".") >= 0){
    console.log("please input a integer");
    pointItemsCount = rowCount-1;
  }
  return pointItemsCount;
}

/**
 * @param originalRange : BodyRange
 * @param pointItemsCount : itemscount that u want
 */
function getPartialRange(originalRange: Excel.Range, pointItemsCount: number ,orientation: number): Excel.Range{
  let partialRange: Excel.Range;
  if(orientation === 1){  //for top n
    partialRange = originalRange.getCell(totalRowCount-pointItemsCount-1,0).getAbsoluteResizedRange(pointItemsCount,1);
  }else{
    partialRange = originalRange.getCell(0,0).getAbsoluteResizedRange(pointItemsCount,1);
  }

  return partialRange;

}

// To calculate the increase for each step between next data list and current data list
//function calculateIncrease(current: Array<Array<number>>, next: Array<Array<number>>, steps: number) {
function calculateIncrease(current: any[][], next: any[][], steps: number): any[] {
  if (current.length != next.length) {
    console.error("Error! current data length:" + current.length + ", next data length" + next.length + ".");
  }

  // let result = new Array(current.length);
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
