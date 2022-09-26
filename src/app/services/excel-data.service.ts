import { Injectable } from "@angular/core";
import { ExportTableDataDto } from "../classes/table-data.class";
import * as XLSX from "xlsx";
import * as _ from "lodash";

@Injectable({
  providedIn: "root",
})
export class ExcelDataService {
  exportDataToExcel(tables: ExportTableDataDto<any>[], fileName: string) {
    if (!tables || (tables && tables.length === 0)) {
      throw new Error("No columns to export");
    }

    const wb = XLSX.utils.book_new();

    for (let i = 0; i < tables.length; i++) {
      const data = [];

      const merges = [];

      const headers = [];
      let mergeAcrossStartC = 0;
      for (let m = 0; m < tables[i].columns.length; m++) {
        merges.push({
          s: { r: 0, c: mergeAcrossStartC },
          e: {
            r: tables[i].columns[m].mergeDown,
            c: mergeAcrossStartC + tables[i].columns[m].mergeAcross,
          },
        });
        mergeAcrossStartC =
          mergeAcrossStartC + tables[i].columns[m].mergeAcross + 1;
        for (let mm = 0; mm <= tables[i].columns[m].mergeAcross; mm++) {
          headers.push(tables[i].columns[m].headerText);
        }
      }
      data.push(headers);

      /**
       * Creating the rows data
       */
      for (let j = 0; j < tables[i].totalRecords; j++) {
        const rowData = [];
        for (let k = 0; k < tables[i].columns.length; k++) {
          if (tables[i].columns[k].mergeAcross > 0) {
            const dataObj = tables[i].columns[k].data[0];
            for (let b = 0; b < tables[i].columns[k].mergeAcross + 1; b++) {
              rowData.push(dataObj[b][j]);
            }
          } else {
            rowData.push(tables[i].columns[k].data[j]);
          }
        }
        data.push(rowData);
      }
      /**
       * Create a worksheet and add data to it
       */
      const ws = XLSX.utils.aoa_to_sheet(data);

      ws["!cols"] = [];
      _.forEach(headers, (val) => {
        ws["!cols"].push({ wpx: 120 });
      });
      ws["!merges"] = merges;
      /**
       * Add worksheet to workbook
       */
      XLSX.utils.book_append_sheet(
        wb,
        ws,
        tables[i].workSheetName || "Sheet" + (i + 1)
      );
    }
    /**
     * Write workbook and force a download
     */
    XLSX.writeFile(wb, fileName + ".xlsx");
  }
}
