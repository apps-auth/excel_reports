part of excel_reports;

class ReportTotalLink {
  final int indexLabel;
  final List<int> indexValues;
  final bool addToTotal;

  ReportTotalLink(this.indexLabel, this.indexValues, {this.addToTotal = true});
}
