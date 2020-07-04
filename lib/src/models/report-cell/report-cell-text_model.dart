part of excel_reports;

class ReportCellText implements IReportCell {
  @override
  CellStyle get cellStyle => style;

  @override
  String get value => text;

  @override
  String get dynamicValue => text;

  final String text;
  final CellStyle style;

  ReportCellText(this.text, {this.style});
}
