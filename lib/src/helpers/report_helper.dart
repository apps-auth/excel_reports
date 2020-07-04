part of excel_reports;

class ReportHelper {
  Future<void> openReport(Excel excel) async {
    Directory appDocDir = await getApplicationDocumentsDirectory();
    excel.encode().then((onValue) async {
      File file = File("${appDocDir.path}/report.xlsx");
      file.createSync(recursive: true);
      file.writeAsBytesSync(onValue);
      OpenFile.open(file.path);
    });
  }

  Excel generateReportHeader(Excel excel, String sheet, ReportHeader header) {
    for (var i = 0; i < header.header.length; i++) {
      ReportCell cell = header.header[i];
      excel.merge(
        sheet,
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i),
        CellIndex.indexByColumnRow(columnIndex: 20, rowIndex: i),
      );
      excel.updateCell(
        sheet,
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i),
        cell.value,
        cellStyle: cell.cellStyle ??
            CellStyle(
              horizontalAlign: HorizontalAlign.Left,
              verticalAlign: VerticalAlign.Center,
            ),
      );
    }
    return excel;
  }

  Excel updateCell({
    @required Excel excel,
    @required String sheet,
    @required dynamic value,
    @required int rowIndex,
    @required int columnIndex,
    CellStyle cellStyle,
  }) {
    excel.updateCell(
      sheet,
      CellIndex.indexByColumnRow(columnIndex: columnIndex, rowIndex: rowIndex),
      value,
      cellStyle: cellStyle ??
          CellStyle(
            horizontalAlign: HorizontalAlign.Center,
            verticalAlign: VerticalAlign.Center,
          ),
    );
    return excel;
  }
}
