part of excel_reports;

class IReportCell {
  final dynamic value;
  final CellStyle cellStyle;
  final dynamic dynamicValue;

  IReportCell(this.value, this.dynamicValue, {this.cellStyle});
}
