part of excel_reports;

class ReportFooter {
  final ReportFooterPosition position;
  final List<ReportRow> rows;

  ReportFooter(
    this.rows, {
    this.position = ReportFooterPosition.EnderStart,
  });
}

enum ReportFooterPosition { EnderStart }
