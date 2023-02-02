import { Injectable } from "@angular/core";
import * as XLSX from "xlsx";
import * as _ from "lodash";

/**
 * Service class to export data to multiple sheets in an excel file.
 * ! Please check SheetJS docs and research before using this code.
 */
@Injectable({
  providedIn: "root",
})
export class ExcelDataService {

  /**
   * Takes data and file to export data to excel file.
   * @param data Multiple tables data will be accepted.
   * @param fileName Name of the excel file.
   */
  exportRawDataToExcel(data: any[], fileName: string) {
    if (!data || (data && data.length === 0)) {
      throw new Error("No data to export");
    }
    const wb = XLSX.utils.book_new();
    data.forEach((value, index) => {
      /**
       * Add worksheet to workbook
       */
      XLSX.utils.book_append_sheet(
        wb,
        this.getWorkSheet(value),
        `Sheet ${index}`
      );
    });
    XLSX.writeFile(wb, fileName + ".xlsx");
  }

  // Create the worksheet from the data passes.
  private getWorkSheet(sheetData: any[]) {
    const data = [];
    const merges = [];
    const headers = [];
    const keys: string[] = _.keys(_.head(sheetData));
    let mergeAcrossStartC = 0;
    for (let m = 0; m < keys.length; m++) {
      merges.push({
        s: { r: 0, c: mergeAcrossStartC },
        e: {
          r: 0,
          c: mergeAcrossStartC,
        },
      });
      mergeAcrossStartC = mergeAcrossStartC + 1;
      headers.push(keys[m]);
    }
    data.push(headers);
    for (let j = 0; j < sheetData.length; j++) {
      const rowData = [];
      for (let k = 0; k < keys.length; k++) {
        rowData.push(sheetData[j][keys[k]]);
      }
      data.push(rowData);
    }
    const ws = XLSX.utils.aoa_to_sheet(data);

    ws["!cols"] = [];
    _.forEach(headers, (val) => {
      ws["!cols"].push({ wpx: 120 });
    });
    ws["!merges"] = merges;
    return ws;
  }
}
