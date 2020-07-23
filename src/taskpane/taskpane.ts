/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Internal used const. DO NOT CHANGE
import * as Ivy from "@ms/charts";

// import {CommonField,BarChartField, ColumnChartField, LineChartField} from "../utils/constants";
// import * as Tool from "../utils/tools";
import { CreateLineChart, PlayLineChart } from "../utils/lineChart";
import { CreateBarChart, PlayBarChart } from "../utils/barChart";
import { CreateColumnChart, PlayColumnChart } from "../utils/columnChart";
import {LineCopyChart, PlayCopyLine} from "../utils/lineCopyChart";
import {DynamicSpace4Line, PlayNewLine} from "../utils/lineDynamicStep";


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

    // new line chart
    document.getElementById("lineCopyChart").onclick = LineCopyChart;
    document.getElementById("playCopyLine").onclick = PlayCopyLine;

    // new line chart
    document.getElementById("dynamicSpace4Line").onclick = DynamicSpace4Line;
    document.getElementById("playNewLine").onclick = PlayNewLine;

  }
});



// let ivyCategoryName = [];
// let ivyValues = [];

// export async function HandleIvy(){
//   try {
//     await Excel.run(async context => {
//       // Find selected table
//       const activeRange = context.workbook.getSelectedRange();
//       let dataTables = activeRange.getTables(false);
//       dataTables.load("items");
//       await context.sync();

//       // Get active table
//       let dataTable = dataTables.items[0];
//       let dataSheet = context.workbook.worksheets.getActiveWorksheet();
//       CommonField.activeTableId = dataTable.id; //id can not be loaded
//       let table = dataSheet.tables.getItem(CommonField.activeTableId);
//       await context.sync();

//       let categoryRange = table.columns.getItemAt(0).getDataBodyRange();
//       let valueRange = table.columns.getItemAt(1).getDataBodyRange();
//       categoryRange.load("values");
//       valueRange.load("values");
//       await context.sync();
//       let tmpCategory = [];
//       let tmpValue = [];
//       for(let i=0;i<categoryRange.values.length;++i){
//         tmpCategory.push(categoryRange.values[i][0]);
//         tmpValue.push(valueRange.values[i][0]);
//       }
//       ivyCategoryName.push(tmpCategory);
//       ivyValues.push(tmpValue);

//     });
//   } catch (error) {
//     console.error(error);
//   }
// }

// const pageRoot = document.getElementById("root");
// pageRoot.style.width = "75%";

// const ivySettings: Ivy.IRenderSettings = {
//   disableChartElementResize: true,
//   renderer: Ivy.IvyRenderer.Svg
// };
// const webLayout: Ivy.ILayoutProvider = new Ivy.EderaLayoutProvider({
//   base: "NOT-SET",
//   layoutUrl: "https://dev.insights.microsoft.com/v3.0/charts?clientId=3D666AF2-59D8-4931-9F29-DDE2E910937B"
// });
// const host = new Ivy.ChartHost(webLayout, ivySettings, pageRoot);

// function layoutChart(): void {
//   const width: number = pageRoot.clientWidth;
//   const chartRatio: number = 0.75;
//   const height: number = chartRatio * width;
//   const testChartSettings: Ivy.IChartSettings = {
//     series: [
//       {
//         data: {
//           categoryNames: [['1','2']],
//           values: [[1,2]]
//         },
//         id: "Series1",
//         layout: "Bar Clustered"  
//       },

//     ],
//     size: {
//       height,
//       width
//     }
//   };
//   host.setConfiguration(testChartSettings);
// }

// (function() {
//   layoutChart();
//   pageRoot.addEventListener("resize", () => layoutChart());
// })();

