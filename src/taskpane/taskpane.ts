/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Internal used const. DO NOT CHANGE
import * as Ivy from "@ms/charts";

import {CommonField,BarChartField, ColumnChartField, LineChartField} from "../utils/constants";
import * as Tool from "../utils/tools";
// import { CreateLineChart, PlayLineChart } from "../utils/lineChart";

// import {LineCopyChart, PlayCopyLine} from "../utils/lineCopyChart";
// import {DynamicSpace4Line, PlayNewLine} from "../utils/lineDynamicStep";

import { CreateBarChart, PlayBarChart } from "../utils/barChart";
import { CreateColumnChart, PlayColumnChart } from "../utils/columnChart";
import { IColor, ISeries } from "@ms/charts";



const pageRoot = document.getElementById("root");
pageRoot.style.width = "75%";

const ivySettings: Ivy.IRenderSettings = {
  disableChartElementResize: true,
  renderer: Ivy.IvyRenderer.Svg
};
const webLayout: Ivy.ILayoutProvider = new Ivy.EderaLayoutProvider({
  base: "NOT-SET",
  layoutUrl: "https://dev.insights.microsoft.com/v3.0/charts?clientId=3D666AF2-59D8-4931-9F29-DDE2E910937B"
});

let ivyCategoryName = []; //use ivyTable to
let ivyValues = [];


Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    //line chart
    // document.getElementById("createLineChart").onclick = CreateLineChart;
    // document.getElementById("playLineChart").onclick = PlayLineChart;
    
    //bar chart
    document.getElementById("createBarChart").onclick = CreateBarChart;
    document.getElementById("playBarChart").onclick = PlayBarChart;
    //column chart
    document.getElementById("createColumnChart").onclick = CreateColumnChart;
    document.getElementById("playColumnChart").onclick = PlayColumnChart;

    // // new line chart
    // document.getElementById("lineCopyChart").onclick = LineCopyChart;
    // document.getElementById("playCopyLine").onclick = PlayCopyLine;

    // // new line chart
    // document.getElementById("dynamicSpace4Line").onclick = DynamicSpace4Line;
    // document.getElementById("playNewLine").onclick = PlayNewLine;

    // slide window line chart
    // document.getElementById("test1").onclick = test1;
    // document.getElementById("test2").onclick = test2;

    //ivychart
    // document.getElementById("test3").onclick = test3;



  }
});


/**
 * slide window
 */
export async function test1() {
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
      CommonField.activeTableId = dataTable.id; //id can not be loaded
      let table = dataSheet.tables.getItem(CommonField.activeTableId);
      await context.sync();

      let wholeRange = table.getRange();
      wholeRange.load("rowCount");
      wholeRange.load("columnCount");
      await context.sync();
      CommonField.totalColumnCount = wholeRange.columnCount;
      CommonField.totalRowCount = wholeRange.rowCount;

      //create toolTable
      //delete the old chart and sheet
      let toolSheet: Excel.Worksheet;
      toolSheet = context.workbook.worksheets.getItemOrNullObject(CommonField.toolSheetName);
      toolSheet.load();
      await context.sync();
      let lastLineChart: Excel.Chart;
      let lastBarChart: Excel.Chart;
      let lastColumnChart: Excel.Chart;

      if (JSON.stringify(toolSheet) !== "{}") {
        lastLineChart = dataSheet.charts.getItemOrNullObject(LineChartField.lineChartName);
        lastBarChart = dataSheet.charts.getItemOrNullObject(BarChartField.barChartName);
        lastColumnChart = dataSheet.charts.getItemOrNullObject(ColumnChartField.columnChartName);
        //chart delete
        lastLineChart.load();
        lastBarChart.load();
        lastColumnChart.load();
        await context.sync();
        if (JSON.stringify(lastLineChart) !== "{}") {
          lastLineChart.delete();
        }
        if (JSON.stringify(lastBarChart) !== "{}") {
          lastBarChart.delete();
        }
        if (JSON.stringify(lastColumnChart) !== "{}") {
          lastColumnChart.delete();
        }
        toolSheet.delete();
      }
      toolSheet = context.workbook.worksheets.add(CommonField.toolSheetName);
      let toolRange = toolSheet
        .getCell(0, 0)
        .getAbsoluteResizedRange(CommonField.totalRowCount, CommonField.totalColumnCount);
      let toolTable = toolSheet.tables.add(toolRange, true);
      toolTable.set({
        name: CommonField.toolTableName
      });

      let toolCategoryRange = toolTable.columns.getItemAt(0).getRange();
      toolCategoryRange.copyFrom(table.columns.getItemAt(0).getRange());
      // let dataRange = dataTable.getRange();
      // toolRange.copyFrom(dataRange);

      let toolColumnCollection = toolTable.columns;
      toolColumnCollection.load("count");
      await context.sync();

      // let toolColumnCount = toolColumnCollection.count;

      // for(let i=1; i< toolTable.columns.count;++i){
      for (let i = 1; i < toolTable.columns.count; ++i) {
        let toolCurRange = toolTable.columns.getItemAt(i).getRange();
        toolCurRange.copyFrom(table.columns.getItemAt(i).getRange());
      }

      //input
      let inputElement = document.getElementById("PointItems") as HTMLInputElement;
      CommonField.inputPointItems = inputElement.value;
      let optionElement = document.getElementById("orientation") as HTMLOptionElement;
      CommonField.orientation = Number(optionElement.value);

      CommonField.pointItemsCount = Tool.formatInput(CommonField.inputPointItems, CommonField.totalRowCount);
      // let initLineRange = Tool.getLinePartialRange(toolTable.getRange(), CommonField.pointItemsCount, toolColumnCount);
      let initLineRange = Tool.getLinePartialRange(toolTable.getRange(), CommonField.pointItemsCount, 4);
      // let initLineRange = Tool.getLinePartialRange(toolTable.getRange(), CommonField.pointItemsCount, 11);

      let initCategoryRange = toolTable
        .getRange()
        .getCell(0, 1)
        .getAbsoluteResizedRange(1, 3);

      if (CommonField.orientation == 1) {
        toolTable.sort.apply([{ key: 1, ascending: false }], true);
      } else {
        toolTable.sort.apply([{ key: 1, ascending: true }], true);
      }
      let lineChart = dataSheet.charts.add(Excel.ChartType.line, initLineRange, "Rows");
      let lineChartHeight = CommonField.chartHeight - 50;
      lineChart.set({
        name: LineChartField.lineChartName,
        height: lineChartHeight,
        width: CommonField.chartWidth,
        left: CommonField.chartLeft,
        top: CommonField.chartTop
      });
      let categoryAxis = lineChart.axes.getItem(Excel.ChartAxisType.category);
      categoryAxis.setCategoryNames(initCategoryRange);

      let setLabel = {
        showSeriesName: true,
        showValue: true,
        numberFormat: "#,##0"
      };
      let unsetLabel = {
        showSeriesName: false,
        showValue: false
      };
      LineChartField.linePointSetLabel = {
        dataLabel: setLabel
      };
      LineChartField.linePointUnsetLabel = {
        dataLabel: unsetLabel
      };

      initLineRange.load("rowCount");
      await context.sync();
      let initRowCount = initLineRange.rowCount;
      console.log(initRowCount);
      // for (let i = initRowCount - 2; i >= 0; --i) {
      //   //getItem(1):The first two columns, only the rightmost column is displayed.
      //   lineChart.series
      //     .getItemAt(i)
      //     .points.getItemAt(1)
      //     .set(LineChartField.linePointSetLabel);
      // }
      lineChart.title.text = "LineChart";
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}


export async function test2() {
  try {
    await Excel.run(async context => {
      let dataSheet = context.workbook.worksheets.getActiveWorksheet();
      // let table = dataSheet.tables.getItem(CommonField.activeTableId);

      //get toolTable
      let toolSheet = context.workbook.worksheets.getItem(CommonField.toolSheetName);
      let toolTable = toolSheet.tables.getItem(CommonField.toolTableName);
      let lineChart = dataSheet.charts.getItem(LineChartField.lineChartName);
      let lineCategoryAxis = lineChart.axes.getItem(Excel.ChartAxisType.category);

      let initialCell = toolTable.getRange().getCell(0, 0);
      initialCell.load();
      await context.sync();

      // for(let columnPointer = 3;columnPointer<CommonField.totalColumnCount;++columnPointer){
      for (let columnPointer = 4; columnPointer < CommonField.totalColumnCount; ++columnPointer) {
        //remove original datalabel
        // for (let j = CommonField.pointItemsCount - 1; j >= 0; --j) {
        //   lineChart.series
        //     .getItemAt(j)
        //     .points.getItemAt(columnPointer - 2)
        //     .set(LineChartField.linePointUnsetLabel); //getItemAt(i - 3): pionts count from 0.
        // }
        // await context.sync();

        // let resizedRange = Tool.getLinePartialRange(initialCell, CommonField.pointItemsCount, columnPointer);
        initialCell = toolTable.getRange().getCell(0, columnPointer-3);
        let resizedRange = Tool.getLinePartialRange(initialCell, CommonField.pointItemsCount, 3);

        resizedRange.load("address");
        await context.sync();
        console.log(resizedRange.address);

        if (CommonField.orientation === 1) {
          toolTable.sort.apply([{ key: columnPointer, ascending: false }], true);
        } else {
          toolTable.sort.apply([{ key: columnPointer, ascending: true }], true);
        }

        lineChart.setData(resizedRange, "Rows"); //dynamic change chart

        //set categoryName
        let resizedNameRange = resizedRange.getCell(0, 1).getAbsoluteResizedRange(1, 3);
        lineCategoryAxis.setCategoryNames(resizedNameRange);

        //set new datalabel
        // for (let j = CommonField.pointItemsCount - 1; j >= 0; --j) {
        //   lineChart.series
        //     .getItemAt(j)
        //     .points.getItemAt(columnPointer - 1)
        //     .set(LineChartField.linePointSetLabel);
        // }
        await context.sync();
        // Tool.sleep(500);
      }
    });
  } catch (error) {
    console.error(error);
  }
}


export async function HandleIvy(){
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
      CommonField.activeTableId = dataTable.id; //id can not be loaded
      let table = dataSheet.tables.getItem(CommonField.activeTableId);
      await context.sync();

      let ivySheet: Excel.Worksheet;
      ivySheet = context.workbook.worksheets.getItemOrNullObject("IvySheet");
      if(JSON.stringify(ivySheet) === "{}"){
        ivySheet.delete();
      }
      await context.sync();

      ivySheet = context.workbook.worksheets.add("IvySheet");
      let ivyRange = ivySheet
        .getCell(0, 0)
        .getAbsoluteResizedRange(11, 11);
      let ivyTable = ivySheet.tables.add(ivyRange, true);
      ivyTable.set({
        name: "ivyTable"
      });
      ivyTable.getRange().copyFrom(table.getRange().getAbsoluteResizedRange(11,11));
      ivyTable.sort.apply([{ key: 1, ascending: true }], true);

      let categoryRange = ivyTable.columns.getItemAt(0).getDataBodyRange();
      let valueRange = ivyTable.columns.getItemAt(1).getDataBodyRange();
      categoryRange.load("values");
      valueRange.load("values");
      await context.sync();
      let tmpCategory = [];
      let tmpValue = [];
      for(let i=0;i<10;++i){
        tmpCategory.push(categoryRange.values[i][0]);
        tmpValue.push(valueRange.values[i][0]);
      }
      ivyCategoryName.push(tmpCategory);
      ivyValues.push(tmpValue);

    });
  } catch (error) {
    console.error(error);
  }
}


const host = new Ivy.ChartHost(webLayout, ivySettings, pageRoot);
//settings: Ivy.IChartSettings
function layoutChart(): void {
  const width: number = pageRoot.clientWidth;

  const chartRatio: number = 0.75;
  const height: number = chartRatio * width;

  let series1: ISeries = {
    data: {
      // categoryNames: [['1','2']],
      // values: [[1,2]]
      categoryNames: ivyCategoryName,
      values: ivyValues
    },
    id: "Series1",
    // layout: "Bar Clustered"
    // layout: "Column Clustered"
    layout: "Line"

  };
  let myColor: IColor = {
    r:255,
    g:0,
    b:139,
    a:0.5
  };
  series1.fillColor = myColor;
  series1.lineColor = myColor;

  const testChartSettings: Ivy.IChartSettings = {
    series: [
      // {
      //   data: {
      //     // categoryNames: [['1','2']],
      //     // values: [[1,2]]
      //     categoryNames: ivyCategoryName,
      //     values: ivyValues
      //   },
      //   id: "Series1",
      //   layout: "Bar Clustered"  
      // },
      series1

    ],
    size: {
      height,
      width
    }
  };
  host.setConfiguration(testChartSettings);
}

export async function test3() {
  await HandleIvy();
  layoutChart();
  //pageRoot.addEventListener("resize", () => layoutChart());
}

