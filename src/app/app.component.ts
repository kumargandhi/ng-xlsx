import { Component } from "@angular/core";
import * as _ from "lodash";
import {
  ExportTableDataColumnDto,
  ExportTableDataDto,
} from "./classes/table-data.class";
import { ExcelDataService } from "./services/excel-data.service";

@Component({
  selector: "app-root",
  templateUrl: "./app.component.html",
  styleUrls: ["./app.component.scss"],
})
export class AppComponent {
  title = "ng-xlsx";

  users = [
    {
      uid: "1",
      first: "Mark",
      last: "Otto",
      handle: "@mdo",
    },
    {
      uid: "2",
      first: "Jacob",
      last: "Thornton",
      handle: "@fat",
    },
    {
      uid: "3",
      first: "Larry the Bird",
      last: "Thornton",
      handle: "@twitter",
    },
  ];

  constructor(private _excelDataService: ExcelDataService) {}

  export() {
    const exportTable: ExportTableDataDto<number | string | {}> =
      new ExportTableDataDto<number | string | {}>();
    exportTable.totalRecords = this.users.length;
    exportTable.workSheetName = `${this.title}`;
    exportTable.columns = [];
    const keys: string[] = _.keys(_.head(this.users));
    keys.forEach((key) => {
      const clmn: ExportTableDataColumnDto<string> =
        new ExportTableDataColumnDto<string>();
      clmn.headerText = key;
      clmn.data = [];
      for (let i = 0; i < this.users.length; i++) {
        clmn.data.push(this.users[i][key]);
      }
      exportTable.columns.push(clmn);
    });
    this._excelDataService.exportDataToExcel([exportTable], `${this.title}`);
  }
}
