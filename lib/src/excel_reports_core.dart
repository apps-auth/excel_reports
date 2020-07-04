part of excel_reports;

class ReportCore {
  Excel generateReport({
    @required Excel excel,
    @required String sheet,
    @required List<ReportDayModel> expenseDays,
    @required CellStyle labelStyle,
    @required CellStyle totalCellStyle,
    @required DateFormat dayFormat,
  }) {
    Map<String, dynamic> map = _generateRow(excel, sheet, expenseDays);
    excel = map['excel'];

    Map<int, Map<String, _ReportCellTotalModel>> totals = map['totals'];
    Map<int, Map<String, _ReportCellTotalModel>> total = map['total'];

    excel = _generateTotalValues(
      excel: excel,
      sheet: sheet,
      total: total,
      totals: totals,
      labelStyle: labelStyle,
      cellStyle: totalCellStyle,
    );

    return excel;
  }

  Map<String, dynamic> _generateRow(
      Excel excel, String sheet, List<ReportDayModel> expenseDays,
      {DateFormat dayFormat}) {
    ReportHelper helper = ReportHelper();
    Sheet sheetObject = excel[sheet];
    int indexLabels = sheetObject.maxRows;

    Map<int, Map<String, _ReportCellTotalModel>> total = {};
    Map<int, Map<String, _ReportCellTotalModel>> totals = {};

    for (var indexDay = 0; indexDay < expenseDays.length; indexDay++) {
      ReportDayModel expenseDay = expenseDays[indexDay];
      DateTime day = expenseDay.day;
      if (indexDay == 0) {
        excel = helper.updateCell(
          excel: excel,
          sheet: sheet,
          value: "Dia",
          cellStyle: expenseDay.expenses.first.expenses.first.label.cellStyle,
          rowIndex: indexLabels,
          columnIndex: 0,
        );
      }

      int rowIndex = sheetObject.maxRows;
      int columnIndex = 0;

      excel = helper.updateCell(
        excel: excel,
        sheet: sheet,
        value: dayFormat ?? DateFormat("dd/MM").format(day),
        cellStyle: expenseDay.expenses.first.expenses.first.cell.cellStyle,
        rowIndex: rowIndex,
        columnIndex: columnIndex,
      );
      columnIndex++;
      for (var expenseIndex = 0;
          expenseIndex < expenseDay.expenses.length;
          expenseIndex++) {
        ReportDayCellModel expense = expenseDay.expenses[expenseIndex];

        if (expense.totalLinks != null && expense.totalLinks.isNotEmpty) {
          expense.totalLinks.forEach((reportTotalLink) {
            ReportDayItenModel cell =
                expense.expenses[reportTotalLink.indexLabel];
            dynamic label = cell.cell.value;
            _ReportCellTotalModel value = _ReportCellTotalModel();

            assert(
                reportTotalLink.indexValues != null &&
                    reportTotalLink.indexValues.isNotEmpty,
                "reportTotalLink.indexValues is not null or empty");

            reportTotalLink.indexValues.forEach((indexValue) {
              IReportCell cell = expense.expenses[indexValue].cell;
              value.type = cell.runtimeType;

              if (cell.runtimeType == ReportCellMoney) {
                if (value.value == null) {
                  value.value = 0;
                }
                value.value += cell.value;
              } else if (cell.runtimeType == ReportCellDuration) {
                if (value.value == null) {
                  value.value = Duration();
                }
                value.value += cell.value;
              }
            });
            if (reportTotalLink.addToTotal) {
              if (total[rowIndex] == null) {
                total[rowIndex] = {label: value.copyWith()};
              } else if (total[rowIndex][label] == null) {
                total[rowIndex][label] = value.copyWith();
              } else {
                total[rowIndex][label].value += value.value;
              }
            }

            if (totals[rowIndex] == null) {
              totals[rowIndex] = {label: value.copyWith()};
            } else if (totals[rowIndex][label] == null) {
              totals[rowIndex][label] = value.copyWith();
            } else {
              totals[rowIndex][label].value += value.value;
            }
          });
        }

        expense.expenses.forEach((cell) {
          if (columnIndex - 1 < sheetObject.maxCols) {
            excel = helper.updateCell(
              excel: excel,
              sheet: sheet,
              value: cell.label.dynamicValue,
              cellStyle: cell.label.cellStyle,
              rowIndex: indexLabels,
              columnIndex: columnIndex,
            );
          }
          excel = helper.updateCell(
            excel: excel,
            sheet: sheet,
            value: cell.cell.dynamicValue,
            cellStyle: cell.cell.cellStyle,
            rowIndex: rowIndex,
            columnIndex: columnIndex,
          );
          columnIndex++;
        });
      }
    }
    // totals.forEach((key, value) {
    //   value.forEach((k, v) {
    //     print("$key: $k ==> ${v.value}");
    //   });
    // });
    return {
      'excel': excel,
      'total': total,
      'totals': totals,
    };
  }

  Excel _generateTotalValues({
    @required Excel excel,
    @required String sheet,
    @required Map<int, Map<String, _ReportCellTotalModel>> total,
    @required Map<int, Map<String, _ReportCellTotalModel>> totals,
    @required CellStyle labelStyle,
    @required CellStyle cellStyle,
  }) {
    ReportHelper helper = ReportHelper();

    Sheet sheetObject = excel[sheet];

    int labelIndex = totals.keys.first - 1;

    int lastColumnIndex = sheetObject.maxCols + 1;
    int lastColumnIndexIncrementable = sheetObject.maxCols + 2;
    Map<int, dynamic> totalOcurenses = {};

    totals.forEach((rowIndex, map) {
      map.forEach((occurence, value) {
        int cIndex;
        List<dynamic> row = sheetObject.rows[labelIndex];

        int i = lastColumnIndex;
        while (cIndex == null && i < row.length) {
          dynamic c = row[i];
          if (c == occurence) {
            cIndex = i;
          }
          i++;
        }

        excel = helper.updateCell(
          excel: excel,
          sheet: sheet,
          value: occurence,
          rowIndex: labelIndex,
          columnIndex: cIndex ?? lastColumnIndexIncrementable,
          cellStyle: labelStyle,
        );
        excel = helper.updateCell(
          excel: excel,
          sheet: sheet,
          value: value.valueFormated,
          rowIndex: rowIndex,
          columnIndex: cIndex ?? lastColumnIndexIncrementable,
        );
        if (cIndex == null) {
          if (totalOcurenses[lastColumnIndexIncrementable] == null) {
            totalOcurenses[lastColumnIndexIncrementable] = value.value;
          } else {
            totalOcurenses[lastColumnIndexIncrementable] += value.value;
          }

          lastColumnIndexIncrementable++;
        } else {
          if (totalOcurenses[cIndex] == null) {
            totalOcurenses[cIndex] = value.value;
          } else {
            totalOcurenses[cIndex] += value.value;
          }
        }
      });
    });
    Type type = total.values.first.values.first.type;

    excel = helper.updateCell(
      excel: excel,
      sheet: sheet,
      value: "Total",
      rowIndex: labelIndex,
      columnIndex: lastColumnIndex,
      cellStyle: labelStyle,
    );
    dynamic totalValue;
    if (type == ReportCellMoney) {
      totalValue = 0;
    } else if (type == ReportCellDuration) {
      totalValue = Duration();
    }

    total.forEach((rowIndex, map) {
      dynamic initialValue;
      if (type == ReportCellMoney) {
        initialValue = 0;
      } else if (type == ReportCellDuration) {
        initialValue = Duration();
      }
      dynamic value = map.values.fold(initialValue,
          (previousValue, element) => previousValue += element.value);
      totalValue += value;

      excel = helper.updateCell(
        excel: excel,
        sheet: sheet,
        value: map.values.first.format(value),
        rowIndex: rowIndex,
        columnIndex: lastColumnIndex,
      );
    });

    int lastRowIndex = sheetObject.maxRows;

    excel = helper.updateCell(
      excel: excel,
      sheet: sheet,
      value: total.values.first.values.first.format(totalValue),
      rowIndex: lastRowIndex,
      // rowIndex: m.length + 1,
      columnIndex: lastColumnIndex,
      cellStyle: cellStyle,
    );

    totalOcurenses.forEach((index, value) {
      excel = helper.updateCell(
        excel: excel,
        sheet: sheet,
        value: total.values.first.values.first.format(value),
        rowIndex: lastRowIndex,
        columnIndex: index,
        cellStyle: cellStyle,
      );
    });

    return excel;
  }
}

class _ReportCellTotalModel {
  dynamic value;
  Type type;
  _ReportCellTotalModel({
    this.value,
    this.type,
  });

  String get valueFormated {
    if (type == ReportCellMoney) {
      final oCcy = NumberFormat.simpleCurrency(locale: "pt_BR");
      return oCcy.format(value);
    } else if (type == ReportCellDuration) {
      return durationToString(value);
    } else {
      return "";
    }
  }

  String format(dynamic v) {
    if (type == ReportCellMoney) {
      final oCcy = NumberFormat.simpleCurrency(locale: "pt_BR");
      return oCcy.format(v);
    } else if (type == ReportCellDuration) {
      return durationToString(v);
    } else {
      return "";
    }
  }

  _ReportCellTotalModel copyWith({
    dynamic value,
    Type type,
  }) {
    return _ReportCellTotalModel(
      value: value ?? this.value,
      type: type ?? this.type,
    );
  }
}
