library excel_reports;

import 'package:flutter/foundation.dart';
import 'package:excel/excel.dart';
import 'package:intl/intl.dart';
import 'package:flutter/material.dart';
import 'dart:io';
import 'package:open_file/open_file.dart';
import 'package:path_provider/path_provider.dart';

part 'src/excel_reports.dart';
part 'src/excel_reports_core.dart';

part 'src/models/report-total-link_model.dart';
part 'src/models/report-header_model.dart';
part 'src/models/report-day_model.dart';
part 'src/models/report-cell_model.dart';

part 'src/models/report-cell/report-cell_interface.dart';
part 'src/models/report-cell/report-cell-duration_model.dart';
part 'src/models/report-cell/report-cell-money_model.dart';
part 'src/models/report-cell/report-cell-text_model.dart';

part 'src/helpers/report_helper.dart';

part 'src/functions/durationToString.dart';

part 'src/enums/report-type_enum.dart';
