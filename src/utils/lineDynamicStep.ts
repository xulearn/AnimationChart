import { BarChartField, ColumnChartField, CommonField, LineChartField } from "./constants";

export async function DynamicSpace4Line() {
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

      toolTable.columns
        .getItemAt(1)
        .getRange()
        .getAbsoluteResizedRange(CommonField.totalRowCount, 3)
        .copyFrom(
          dataTable.columns
            .getItemAt(1)
            .getRange()
            .getAbsoluteResizedRange(CommonField.totalRowCount, 3)
        );

      let initLineRange = toolTable.columns
        .getItemAt(0)
        .getRange()
        .getAbsoluteResizedRange(20, 4);
      let lineChart = dataSheet.charts.add(Excel.ChartType.line, initLineRange, "Rows");

      let initCategoryRange = toolTable.columns
        .getItemAt(0)
        .getRange()
        .getCell(0, 1)
        .getAbsoluteResizedRange(1, 3);
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
    });
  } catch (error) {
    console.error(error);
  }
}

export async function PlayNewLine() {
  try {
    await Excel.run(async context => {
      let dataSheet = context.workbook.worksheets.getActiveWorksheet();
      let dataTable = dataSheet.tables.getItem(CommonField.activeTableId);

      //get toolTable
      let toolSheet = context.workbook.worksheets.getItem(CommonField.toolSheetName);
      let toolTable = toolSheet.tables.getItem(CommonField.toolTableName);
      let lineChart = dataSheet.charts.getItem(LineChartField.lineChartName);

      let i = 0;
      let step = 1;
      let threshold = 10;
      let j = 0;
      let resizedRange: Excel.Range;
      while (j <= CommonField.totalColumnCount) {
        j = i * step + 1;

        let newRange: Excel.Range = toolTable.columns.getItemAt(i + 1).getRange();
        let dataRange: Excel.Range = dataTable.columns.getItemAt(j).getRange();

        newRange.copyFrom(dataRange);

        if (i >= 2) {
          resizedRange = toolTable.getRange().getAbsoluteResizedRange(CommonField.totalRowCount, i + 2);
          lineChart.setData(resizedRange, "Rows");
        }

        if (CommonField.totalColumnCount - 1 - j <= step) {
          //the next j + step > totalColumnCount
          newRange = toolTable.columns.getItemAt(i + 2).getRange();
          dataRange = dataTable.columns.getItemAt(CommonField.totalColumnCount - 1).getRange();
          newRange.copyFrom(dataRange);

          resizedRange = toolTable.getRange().getAbsoluteResizedRange(CommonField.totalRowCount, i + 3);
          lineChart.setData(resizedRange, "Rows");
          break;
        }

        await context.sync();
        i++;
        if (i === threshold) {
          i = 0;
          step += 1;
        }
      }
    });
  } catch (error) {
    console.error(error);
  }
}
