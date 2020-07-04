part of excel_reports;

class ReportCellMoney implements IReportCell {
  @override
  CellStyle get cellStyle => style;

  @override
  double get value => moneyValue;

  @override
  String get dynamicValue => numberFormat.format(moneyValue);

  final NumberFormat numberFormat;
  final double moneyValue;
  final CellStyle style;

  ReportCellMoney(this.numberFormat, this.moneyValue, {this.style});
}
