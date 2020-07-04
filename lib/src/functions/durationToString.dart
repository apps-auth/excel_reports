part of excel_reports;

String durationToString(Duration duration) {
  int hour = _getDurationHour(duration);
  int minute = _getDurationMinute(duration);
  return "${_correctFormatTime(hour)}:${_correctFormatTime(minute)}";
}

int _getDurationHour(Duration duration) {
  int hour = 0;
  if (duration.toString().length == 16) {
    hour = int.parse(duration.toString().substring(0, 3).toString());
  } else if (duration.toString().length == 14) {
    hour = int.parse(duration.toString().substring(0, 1).toString());
  } else {
    hour = int.parse(duration.toString().substring(0, 2).toString());
  }
  return hour;
}

int _getDurationMinute(Duration duration) {
  int minute = 0;
  if (duration.toString().length == 14) {
    minute = int.parse(duration.toString().substring(2, 4).toString());
  } else {
    minute = int.parse(duration.toString().substring(3, 5).toString());
  }
  return minute;
}

String _correctFormatTime(int time) {
  String t = time.toString();
  if (t.length < 2) {
    t = "0$t";
  }
  return t;
}
