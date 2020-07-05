part of excel_reports;

class Report {
  Report();

  Excel excel;

  Excel generateReport({
    ReportHeader header,
    List<ReportDayModel> reportDays,
    CellStyle labelStyle,
    CellStyle totalCellStyle,
    DateFormat dayFormat,
    ReportFooter footer,
  }) {
    assert(reportDays != null && reportDays.isNotEmpty,
        "List<ExpenseDayModel> expenseDays is require data on type ReportType.Expense");
    excel = _generateReportExpense(
      header: header,
      expenseDays: reportDays,
      labelStyle: labelStyle,
      totalCellStyle: totalCellStyle,
      dayFormat: dayFormat,
      footer: footer,
    );
    return excel;
  }

  Excel _generateReportExpense({
    @required ReportHeader header,
    @required List<ReportDayModel> expenseDays,
    @required CellStyle labelStyle,
    @required CellStyle totalCellStyle,
    DateFormat dayFormat,
    ReportFooter footer,
  }) {
    ReportCore reportExpense = ReportCore();
    ReportHelper helper = ReportHelper();
    excel = Excel.createExcel();
    String sheet = 'Sheet1';

    if (header != null) {
      excel = helper.generateReportHeader(excel, sheet, header);
    }

    excel = reportExpense.generateReport(
      excel: excel,
      sheet: sheet,
      expenseDays: expenseDays,
      labelStyle: labelStyle,
      totalCellStyle: totalCellStyle,
      dayFormat: dayFormat,
      footer: footer,
    );

    return excel;
  }

  Future<void> openReport() async {
    ReportHelper helper = ReportHelper();
    await helper.openReport(excel);
    return;
  }
}
