import { CommonField, LineChartField } from "./constants";

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
      CommonField.activeTableId = dataTable.id;
      let table = dataSheet.tables.getItem(CommonField.activeTableId);
      await context.sync();

      let wholeRange = table.getRange();
      wholeRange.load("rowCount");
      wholeRange.load("columnCount");
      await context.sync();

      CommonField.totalColumnCount = wholeRange.columnCount;
      CommonField.totalRowCount = wholeRange.rowCount;

      //get initial range, at least one category column and two data columns
      let initialCell = table.getRange().getCell(0, 0);
      let initialRange = initialCell.getAbsoluteResizedRange(CommonField.totalRowCount, 3); //It's not an offset, it's an absolute amount .
      //if the category sorted by columns
      let initalCategoryName = table
        .getRange()
        .getCell(0, 1)
        .getAbsoluteResizedRange(1, 2);

      let lineChart = dataSheet.charts.add(Excel.ChartType.line, initialRange, "Rows");
      // lineChart.set({
      //   name: lineChartName
      // });
      let lineChartHeight = CommonField.chartHeight - 50;
      lineChart.set({
        name: LineChartField.lineChartName,
        height: lineChartHeight,
        width: CommonField.chartWidth,
        left: CommonField.chartLeft,
        top: CommonField.chartTop
      });
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
      LineChartField.linePointSetLabel = {
        dataLabel: setLabel
      };
      LineChartField.linePointUnsetLabel = {
        dataLabel: unsetLabel
      };
      for (let i = CommonField.totalRowCount - 2; i >= 0; --i) {
        //getItem(1):The first two columns, only the rightmost column is displayed.
        //totalRowCount - 2:series counts from 0, totalRowCount counts from 1.
        lineChart.series
          .getItemAt(i)
          .points.getItemAt(1)
          .set(LineChartField.linePointSetLabel);
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
      let table = dataSheet.tables.getItem(CommonField.activeTableId);
      let lineChart = dataSheet.charts.getItem(LineChartField.lineChartName);
      let categoryAxis = lineChart.axes.getItem(Excel.ChartAxisType.category);
      await context.sync();

      //get initial range, at least one category column and two data columns
      let initialCell = table.getRange().getCell(0, 0);
      let initialRange = initialCell.getAbsoluteResizedRange(CommonField.totalRowCount, 3); //It's not an offset, it's an absolute amount .
      let initalCategoryName = table
        .getRange()
        .getCell(0, 1)
        .getAbsoluteResizedRange(1, 2);

      for (let i = 4; i <= CommonField.totalColumnCount; ++i) {
        // 4th column
        //todo!!
        //from the forth column

        let resizedRange = initialRange.getAbsoluteResizedRange(CommonField.totalRowCount, i); //Translate one column unit to the right.

        //todo: do not use sleep
        // sleep(1000);
        lineChart.setData(resizedRange, "Rows"); //dynamic change chart

        //set categoryName
        let resizedNameRange = initalCategoryName.getAbsoluteResizedRange(1, i);
        categoryAxis.setCategoryNames(resizedNameRange);

        //remove original datalabel
        for (let j = CommonField.totalRowCount - 2; j >= 0; --j) {
          lineChart.series
            .getItemAt(j)
            .points.getItemAt(i - 3)
            .set(LineChartField.linePointUnsetLabel); //getItemAt(i - 3): pionts count from 0.
        }
        await context.sync();

        //set new datalabel
        for (let j = CommonField.totalRowCount - 2; j >= 0; --j) {
          lineChart.series
            .getItemAt(j)
            .points.getItemAt(i - 2)
            .set(LineChartField.linePointSetLabel);
        }
        await context.sync();
      }
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
