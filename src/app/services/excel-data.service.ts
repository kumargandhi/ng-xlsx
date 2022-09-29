import { Injectable } from "@angular/core";
import * as XLSX from "xlsx";
import * as _ from "lodash";

@Injectable({
  providedIn: "root",
})
export class ExcelDataService {
  exportRawDataToExcel(tableData: any, fileName: string) {
    if (!tableData || (tableData && tableData.length === 0)) {
      throw new Error("No data to export");
    }
    const wb = XLSX.utils.book_new();
    const data = [];
    const merges = [];
    const headers = [];
    const keys: string[] = _.keys(_.head(tableData));
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
    for (let j = 0; j < tableData.length; j++) {
      const rowData = [];
      for (let k = 0; k < keys.length; k++) {
        rowData.push(tableData[j][keys[k]]);
      }
      data.push(rowData);
    }
    const ws = XLSX.utils.aoa_to_sheet(data);

    ws["!cols"] = [];
    _.forEach(headers, (val) => {
      ws["!cols"].push({ wpx: 120 });
    });
    ws["!merges"] = merges;
    /**
     * Add worksheet to workbook
     */
    XLSX.utils.book_append_sheet(wb, ws, "Sheet 1");
    XLSX.writeFile(wb, fileName + ".xlsx");
  }
}
