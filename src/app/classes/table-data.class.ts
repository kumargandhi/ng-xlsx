export class ExportTableDataDto<T> {
  title: string;

  columns: ExportTableDataColumnDto<T>[];

  workSheetName: string;

  totalRecords: number;
}

export class ExportTableDataColumnDto<T> {
  headerText: string;

  data: T[];

  mergeDown = 0;

  mergeAcross = 0;
}
