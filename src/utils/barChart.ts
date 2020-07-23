import { BarChartField } from "./constants";
import { CreateBarOrColumnChart, PlayBarOrColumnChart } from "./rectangleChart";

/**
 * create for bar chart
 */
export async function CreateBarChart() {
  try {
    await CreateBarOrColumnChart(BarChartField.barChartFlag);
  } catch (error) {
    console.error(error);
  }
}

/**
 * play for bar chart
 */
export async function PlayBarChart() {
  try {
    await PlayBarOrColumnChart(BarChartField.barChartFlag);
  } catch (error) {
    console.error(error);
  }
}
