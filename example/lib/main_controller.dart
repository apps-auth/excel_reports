import 'package:excel_reports/excel_reports.dart';
import 'package:excel/excel.dart';

class MainController {
  Future<void> generateExcelExpenses() async {
    Report report = Report();

    List<ReportDayModel> reportExpense = <ReportDayModel>[];

    List<ReportDayItenModel> expenses = <ReportDayItenModel>[];
    List<ReportDayItenModel> expenses2 = <ReportDayItenModel>[];

    CellStyle labeltyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      backgroundColorHex: "#9BC2E6",
    );

    expenses.add(ReportDayItenModel(
        ReportCellText("Ocorrência", style: labeltyle),
        ReportCellText("Refeição")));
    expenses.add(ReportDayItenModel(
        ReportCellText("Horário", style: labeltyle), ReportCellText("12:31")));
    expenses.add(ReportDayItenModel(
        ReportCellText("Comprovante", style: labeltyle), ReportCellText("")));
    expenses.add(ReportDayItenModel(
        ReportCellText("Valor gasto", style: labeltyle),
        ReportCellDuration(
          Duration(minutes: 25),
          style: CellStyle(
            horizontalAlign: HorizontalAlign.Center,
            verticalAlign: VerticalAlign.Center,
            backgroundColorHex: "#00E700",
          ),
        )));

    expenses2.add(ReportDayItenModel(
        ReportCellText("Ocorrência", style: labeltyle),
        ReportCellText("Mecânica")));
    expenses2.add(ReportDayItenModel(
        ReportCellText("Horário", style: labeltyle), ReportCellText("12:31")));
    expenses2.add(ReportDayItenModel(
        ReportCellText("Comprovante", style: labeltyle), ReportCellText("")));
    expenses2.add(ReportDayItenModel(
        ReportCellText("Valor gasto", style: labeltyle),
        ReportCellDuration(
          Duration(minutes: 25),
          style: CellStyle(
            horizontalAlign: HorizontalAlign.Center,
            verticalAlign: VerticalAlign.Center,
            backgroundColorHex: "#00E700",
          ),
        )));
    ReportDayCellModel expenseDay =
        ReportDayCellModel(expenses, totalLinks: <ReportTotalLink>[
      ReportTotalLink(0, [3]),
      ReportTotalLink(1, [3], addToTotal: false),
    ]);
    ReportDayCellModel expenseDay2 =
        ReportDayCellModel(expenses2, totalLinks: <ReportTotalLink>[
      ReportTotalLink(0, [3]),
      ReportTotalLink(1, [3], addToTotal: false),
    ]);

    reportExpense.add(ReportDayModel(DateTime.now(), [
      expenseDay,
      expenseDay2,
      expenseDay,
    ]));

    reportExpense.add(ReportDayModel(DateTime.now().add(Duration(days: 1)), [
      expenseDay,
      expenseDay2,
      expenseDay,
    ]));

    ReportHeader header = ReportHeader(<ReportCell>[
      ReportCell("Relatório de despesas: Schemaq"),
      ReportCell("Motorista: Adriel Auth"),
      ReportCell("Período: 22/06/2020 à 25/06/2020"),
    ]);
    Excel excel = report.generateReport(
      header: header,
      reportDays: reportExpense,
      labelStyle: labeltyle,
      totalCellStyle: CellStyle(
        horizontalAlign: HorizontalAlign.Center,
        verticalAlign: VerticalAlign.Center,
        backgroundColorHex: "#A9D08E",
      ),
    );

    report.openReport();
  }
}
