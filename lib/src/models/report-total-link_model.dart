part of excel_reports;

class ReportTotalLink {
  final int indexLabel;
  final List<int> indexValues;
  final bool addToTotal;
  final bool Function(IReportCell) condition;
  final String title;

  ReportTotalLink(
    this.indexLabel,
    this.indexValues, {
    this.addToTotal = true,
    this.condition,
    this.title,
  });
}
