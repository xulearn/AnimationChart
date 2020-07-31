export module CommonField {
  export let activeTableId;
  export let inputPointItems: any;
  export let orientation: number;

  export let pointItemsCount: number;

  //for original table
  export let totalColumnCount: number;
  export let totalRowCount: number;

  //--------------------------------
  // Parameters. Modify it if needed.
  export const chartWidth = 750,
    chartHeight = 550,
    chartLeft = 150,
    chartTop = 50;

  export const splitIncreasement = 2;

  export const colorList = [
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
  export const fontSize_Title = 28,
    fontSize_CategoryName = 13,
    fontSize_AxisValue = 11,
    fontSize_DataLabel = 13;

  // Internal used const. DO NOT CHANGE
  //for barChart and columnChart and tool
  export const toolSheetName = "toolSheet9527"; //+UUID
  export const toolTableName = "toolTable9527"; //+UUID
}

export module LineChartField {
  //for line chart
  export let lineChartName = "LineChartName9527";
  export let linePointSetLabel;
  export let linePointUnsetLabel;
  // let seriesUpdate;
}

export module BarChartField {
  export const barChartName = "BarChartName9527";
  export const barChartFlag = 1;
}

export module ColumnChartField {
  export const columnChartName = "ColumnChartName9527";
  export const columnChartFlag = 2;
}

export module ToolTableField {
  export enum ToolTableColumnIndex {
    category = 0,
    value = 1,
    color = 2,
    map = 3
  }

  export enum ToolTableColumnName {
    category = "categoryColumn",
    //value
    color = "colorColumn",
    map = "mapColumn"
  }
}
