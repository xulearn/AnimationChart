import { CommonField } from "./constants";

export function formatInput(input: string, rowCount: number): number {
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
export function getPartialRange(originalRange: Excel.Range, pointItemsCount: number, orientation: number): Excel.Range {
  let partialRange: Excel.Range;
  if (orientation === 1) {
    //for top n
    partialRange = originalRange
      .getCell(CommonField.totalRowCount - pointItemsCount - 1, 0)
      .getAbsoluteResizedRange(pointItemsCount, 1);
  } else {
    partialRange = originalRange.getCell(0, 0).getAbsoluteResizedRange(pointItemsCount, 1);
  }

  return partialRange;
}

/**
 * @param originalRange : BodyRange
 * @param pointItemsCount : itemscount that u want
 */
export function getLinePartialRange(originalCell: Excel.Range, pointItemsCount: number, columnsCount:number): Excel.Range {
  let partialRange: Excel.Range;
  partialRange = originalCell.getAbsoluteResizedRange(pointItemsCount+1, columnsCount);
  return partialRange;
}

// To calculate the increase for each step between next data list and current data list
//function calculateIncrease(current: Array<Array<number>>, next: Array<Array<number>>, steps: number) {
export function calculateIncrease(current: any[][], next: any[][], steps: number): any[] {
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

export function mapTargetRangeValue(mapRange: Excel.Range, targetRange: Excel.Range): any[][] {
  let targetArr = [];
  let mapArr = mapRange.values;
  for (let j = 0; j < mapArr.length; ++j) {
    let mapIndex = mapArr[j][0];
    let mapVal = targetRange.values[mapIndex][0];
    targetArr.push([mapVal]);
  }
  return targetArr;
}

export function hiddenSheet(sheet: Excel.Worksheet):void {
  sheet.set({ visibility: "Hidden" });
  // sheet.set({ visibility: "Visible"});
}

export function sleep(sleepTime: number):void {
  var start = new Date().getTime();
  for (var i = 0; i < 1e7; i++) {
    if (new Date().getTime() - start > sleepTime) {
      break;
    }
  }
}
