part of excel_reports;

class ReportDayModel {
  final DateTime day;
  final List<ReportDayCellModel> expenses;

  ReportDayModel(this.day, this.expenses);
}

class ReportDayCellModel {
  final List<ReportDayItenModel> expenses;
  final List<ReportTotalLink> totalLinks;

  ReportDayCellModel(this.expenses, {this.totalLinks});
}

class ReportDayItenModel {
  final IReportCell label;
  final IReportCell cell;

  ReportDayItenModel(this.label, this.cell);
}
