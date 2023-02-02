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

  customers = [
    {
      uid: "1",
      first: "Johnson",
      last: "White",
      handle: "@mdo",
    },
    {
      uid: "2",
      first: "Ashley",
      last: "Borden",
      handle: "@fat",
    },
    {
      uid: "3",
      first: "Marjorie",
      last: "Green",
      handle: "@twitter",
    },
  ];

  selectedTab = "Users";

  constructor(private _excelDataService: ExcelDataService) {}

  exportRawData() {
    this._excelDataService.exportRawDataToExcel([this.users, this.customers], this.title);
  }
}
