import { ColumnChartField } from './constants';
import { CreateBarOrColumnChart, PlayBarOrColumnChart } from './rectangleChart';

/**
 * create for column chart
 */
export async function CreateColumnChart() {
  try {
    await CreateBarOrColumnChart(ColumnChartField.columnChartFlag);
  } catch (error) {
    console.error(error);
  }
}

/**
 * play for column chart
 */
export async function PlayColumnChart() {
  try {
    await PlayBarOrColumnChart(ColumnChartField.columnChartFlag);
  } catch (error) {
    console.error(error);
  }
}
