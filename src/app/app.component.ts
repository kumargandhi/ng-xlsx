import { Component } from "@angular/core";
import * as _ from "lodash";
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

  exportRawData() {
    this._excelDataService.exportRawDataToExcel(this.users, this.title);
  }
}
