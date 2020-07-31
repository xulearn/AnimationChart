import { CommonField, BarChartField, ColumnChartField, ToolTableField } from "./constants";
import { Country } from "./country";
import * as Tool from "./tools";

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
      let lastBarChart: Excel.Chart;
      let lastColumnChart: Excel.Chart;

      if (JSON.stringify(toolSheet) !== "{}") {
        lastBarChart = dataSheet.charts.getItemOrNullObject(BarChartField.barChartName);
        lastColumnChart = dataSheet.charts.getItemOrNullObject(ColumnChartField.columnChartName);
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
      toolSheet = context.workbook.worksheets.add(CommonField.toolSheetName);

      Tool.hiddenSheet(toolSheet);
      await context.sync();

      let toolRange = toolSheet.getCell(0, 0).getAbsoluteResizedRange(CommonField.totalRowCount, 4);
      let toolTable = toolSheet.tables.add(toolRange, true);
      toolTable.set({
        name: CommonField.toolTableName
      });
      //set columnName
      toolTable.columns
        .getItemAt(ToolTableField.ToolTableColumnIndex.category)
        .set({ name: ToolTableField.ToolTableColumnName.category });
      toolTable.columns
        .getItemAt(ToolTableField.ToolTableColumnIndex.color)
        .set({ name: ToolTableField.ToolTableColumnName.color });
      toolTable.columns
        .getItemAt(ToolTableField.ToolTableColumnIndex.map)
        .set({ name: ToolTableField.ToolTableColumnName.map });

      let categoryBodyRange = toolTable.columns.getItem(ToolTableField.ToolTableColumnName.category).getDataBodyRange();
      let curIteratedRange = toolTable.columns.getItemAt(ToolTableField.ToolTableColumnIndex.value).getRange();
      let curIteratedBodyRange = toolTable.columns
        .getItemAt(ToolTableField.ToolTableColumnIndex.value)
        .getDataBodyRange();
      let colorBodyRange = toolTable.columns.getItem(ToolTableField.ToolTableColumnName.color).getDataBodyRange();
      let mapBodyRange = toolTable.columns.getItem(ToolTableField.ToolTableColumnName.map).getDataBodyRange();

      //copy Range
      categoryBodyRange.copyFrom(table.columns.getItemAt(0).getDataBodyRange());
      curIteratedRange.copyFrom(table.columns.getItemAt(1).getRange()); //copy headers too

      colorBodyRange.load("values");
      await context.sync();
      let tmpColorArr = [];
      for (let i = 0; i < CommonField.totalRowCount - 1; ++i) {
        tmpColorArr.push([CommonField.colorList[i % CommonField.colorList.length]]);
      }
      colorBodyRange.values = tmpColorArr;

      mapBodyRange.load("values");
      await context.sync();
      let tmpMapArr = [];
      for (let i = 1; i < CommonField.totalRowCount; ++i) {
        tmpMapArr.push([i]);
      }
      mapBodyRange.values = tmpMapArr;

      //input
      let inputElement = document.getElementById("PointItems") as HTMLInputElement;
      CommonField.inputPointItems = inputElement.value;
      let optionElement = document.getElementById("orientation") as HTMLOptionElement;
      CommonField.orientation = Number(optionElement.value);

      //get input and target items
      CommonField.pointItemsCount = Tool.formatInput(CommonField.inputPointItems, CommonField.totalRowCount);

      let targetIteratedBodyRange = Tool.getPartialRange(
        curIteratedBodyRange,
        CommonField.pointItemsCount,
        CommonField.orientation
      );
      let targetCategoryRange = Tool.getPartialRange(
        categoryBodyRange,
        CommonField.pointItemsCount,
        CommonField.orientation
      );

      // Create Chart
      toolTable.sort.apply([{ key: 1, ascending: true }], true); //toolTable only does ascending sort
      let chart: Excel.Chart;
      if (flag === BarChartField.barChartFlag) {
        chart = dataSheet.charts.add(Excel.ChartType.barClustered, targetIteratedBodyRange);
        chart.set({
          name: BarChartField.barChartName,
          height: CommonField.chartHeight,
          width: CommonField.chartWidth,
          left: CommonField.chartLeft,
          top: CommonField.chartTop
        });
      } else {
        chart = dataSheet.charts.add(Excel.ChartType.columnClustered, targetIteratedBodyRange);
        chart.set({
          name: ColumnChartField.columnChartName,
          height: CommonField.chartHeight,
          width: CommonField.chartWidth,
          left: CommonField.chartLeft,
          top: CommonField.chartTop
        });
      }

      let curheaderRange = curIteratedRange.getCell(0, 0);
      curheaderRange.load("text");
      await context.sync();
      // Set chart tile and style
      chart.title.text = curheaderRange.text[0][0];
      chart.title.format.font.set({ size: CommonField.fontSize_Title });
      chart.legend.set({ visible: false });

      // Set Axis
      let categoryAxis = chart.axes.getItem(Excel.ChartAxisType.category);
      categoryAxis.setCategoryNames(targetCategoryRange);
      categoryAxis.set({ visible: true });
      categoryAxis.format.font.set({ size: CommonField.fontSize_CategoryName });
      let valueAxis = chart.axes.getItem(Excel.ChartAxisType.value);
      valueAxis.format.font.set({ size: CommonField.fontSize_AxisValue });

      let series = chart.series.getItemAt(0);
      series.set({ hasDataLabels: true, gapWidth: 30 });
      series.dataLabels.set({ showCategoryName: false, numberFormat: "#,##0" });
      series.dataLabels.format.font.set({ size: CommonField.fontSize_DataLabel });
      series.points.load();
      await context.sync();

      colorBodyRange.load("values");
      await context.sync();
      let sortedColorArr = colorBodyRange.values;

      // Set data points color
      for (let i = 0; i < series.points.count; i++) {
        if (CommonField.orientation === 1) {
          series.points
            .getItemAt(i)
            .format.fill.setSolidColor(
              sortedColorArr[CommonField.totalRowCount - CommonField.pointItemsCount - 1 + i][0]
            );
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
 * play bar or column chart
 */
export async function PlayBarOrColumnChart(flag: number) {
  try {
    await Excel.run(async context => {
      console.log(CommonField.inputPointItems);
      console.log(CommonField.orientation);
      console.log(CommonField.pointItemsCount);

      let dataSheet = context.workbook.worksheets.getActiveWorksheet();
      let table = dataSheet.tables.getItem(CommonField.activeTableId);

      //get toolTable
      let toolSheet = context.workbook.worksheets.getItem(CommonField.toolSheetName);
      // let toolTable = dataSheet.tables.getItem(toolTableName);
      let toolTable = toolSheet.tables.getItem(CommonField.toolTableName);

      let categoryBodyRange = toolTable.columns
        .getItemAt(ToolTableField.ToolTableColumnIndex.category)
        .getDataBodyRange();
      let curIteratedHeaderRange = toolTable.columns
        .getItemAt(ToolTableField.ToolTableColumnIndex.value)
        .getHeaderRowRange();
      let curIteratedBodyRange = toolTable.columns
        .getItemAt(ToolTableField.ToolTableColumnIndex.value)
        .getDataBodyRange();
      let mapBodyRange = toolTable.columns.getItem(ToolTableField.ToolTableColumnName.map).getDataBodyRange();
      let colorBodyRange = toolTable.columns.getItem(ToolTableField.ToolTableColumnName.color).getDataBodyRange();

      let chart: Excel.Chart;
      if (flag == BarChartField.barChartFlag) {
        chart = dataSheet.charts.getItem(BarChartField.barChartName);
      } else {
        chart = dataSheet.charts.getItem(ColumnChartField.columnChartName);
      }
      //todo splitIncreasement input

      categoryBodyRange.load("values");
      curIteratedBodyRange.load("values");
      mapBodyRange.load("values");
      colorBodyRange.load("values");
      await context.sync();

      //initial countryArr
      let countryArray: Country[] = [];
      for (let i = 0; i < CommonField.totalRowCount - 1; ++i) {
        let curCategory = categoryBodyRange.values[i][0];
        let curValue = curIteratedBodyRange.values[i][0];
        let curMap = mapBodyRange.values[i][0];
        let curColor = colorBodyRange.values[i][0];

        let curCountry = new Country(curCategory, curValue, curMap, 0, curColor);
        countryArray.push(curCountry);
      }

      // paly
      for (let i = 2; i < CommonField.totalColumnCount; ++i) {
        let nextIteratedHeaderRange = table.columns.getItemAt(i).getHeaderRowRange(); //from table
        let nextIteratedRange = table.columns.getItemAt(i).getRange();

        nextIteratedRange.load("values");
        curIteratedBodyRange.load("values");
        curIteratedHeaderRange.load("text");
        mapBodyRange.load("values");
        await context.sync();

        let nextArr = Tool.mapTargetRangeValue(mapBodyRange, nextIteratedRange);
        // Calculate increase based on current value and next value
        let increaseData = Tool.calculateIncrease(curIteratedBodyRange.values, nextArr, CommonField.splitIncreasement);

        for (let j = 0; j < CommonField.totalRowCount - 1; ++j) {
          countryArray[j].setIncreasement(increaseData[j]);
        }

        for (let step = 1; step <= CommonField.splitIncreasement; step++) {
          if (step === CommonField.splitIncreasement) {
            mapBodyRange.load("values");
            await context.sync();
            //The mapRange here is the one that was ordered in the previous 'else', and you'll have to take it again because countryArr already sorted.
            nextArr = Tool.mapTargetRangeValue(mapBodyRange, nextIteratedRange);

            for (let j = 0; j < CommonField.totalRowCount - 1; ++j) {
              countryArray[j].setValue(nextArr[j][0]);
            }
            //set title
            curIteratedHeaderRange.copyFrom(nextIteratedHeaderRange);
          } else {
            // Add increase amount
            for (let j = 0; j < CommonField.totalRowCount - 1; ++j) {
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
          for (let j = 0; j < CommonField.totalRowCount - 1; ++j) {
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
            if (CommonField.orientation === 1) {
              series.points
                .getItemAt(k)
                .format.fill.setSolidColor(
                  tmpColorArr[CommonField.totalRowCount - CommonField.pointItemsCount - 1 + k][0]
                );
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
