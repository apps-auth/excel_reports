part of excel_reports;

class ReportCellDuration implements IReportCell {
  @override
  CellStyle get cellStyle => style;

  @override
  Duration get value => durationValue;

  @override
  String get dynamicValue => durationToString(durationValue);

  final Duration durationValue;
  final CellStyle style;

  ReportCellDuration(this.durationValue, {this.style});
}
