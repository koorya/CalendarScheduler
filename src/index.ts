const SETTINGS_SHEET_NAME = 'Настройки';
const SCHEDULE_SHEET_NAME = 'Расписание';
const CALENDAR_ID_PROP_NAME = 'CALENDAR_ID';
const ACTIVE_COLOR = '#ff0000';
const INACTIVE_COLOR = '#ffffff';

function main() {
  const { days_back, days_fw } = getSettings();
  const calendar_events = getCalendarEvents(days_back, days_fw);

  const display_sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCHEDULE_SHEET_NAME);
  if (!display_sheet) throw Error('sheet not found');

  clearScreen(display_sheet, days_back + days_fw + 1);
  renderHeader(display_sheet);
  renderDayGrid({ days_fw, days_back, display_sheet });

  renderPlaces(calendar_events, display_sheet);

  renderEvents({
    day_offset: days_back,
    days_fw,
    x_offset: 2,
    y_offset: 4,
    calendar_events,
    display_sheet,
  });
}

function renderHeader(display_sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  display_sheet
    .getRange(1, 3, 1, display_sheet.getMaxColumns())
    .breakApart()
    .merge();
}

function renderPlaces(
  calendar_events: GoogleAppsScript.Calendar.CalendarEvent[],
  display_sheet: GoogleAppsScript.Spreadsheet.Sheet
) {
  const places = calendar_events
    .filter((e) => e.getStartTime().getHours() == 5)
    .sort((a, b) => a.getStartTime().getTime() - b.getStartTime().getTime())
    .map((e) => e.getTitle());

  const edges: { start: number; title: string; length: number }[] = [];
  places.forEach((title, idx) => {
    const prev = edges.length
      ? edges[edges.length - 1]
      : { start: 0, length: 0, title: '' };
    if (idx == places.length - 1) {
      const length = idx - prev.start - prev.length + 1;
      const start = prev.start + prev.length;
      edges.push({
        start,
        length,
        title,
      });
    } else if (places[idx + 1] != title) {
      const length = idx - prev.start - prev.length + 1;
      const start = prev.start + prev.length;
      edges.push({
        start,
        length,
        title,
      });
    }
  });

  display_sheet.getRange(4, 3, 1, display_sheet.getMaxColumns()).breakApart();

  edges.map(({ start, length, title }) =>
    display_sheet
      .getRange(4, start + 3, 1, length)
      .merge()
      .setValue(title.toUpperCase())
      .setHorizontalAlignment('center')
  );
}

function renderDayGrid(props: {
  days_fw: number;
  days_back: number;
  display_sheet: GoogleAppsScript.Spreadsheet.Sheet;
}) {
  const { days_back, days_fw, display_sheet } = props;
  Array.from({ length: days_fw + days_back + 1 }, (x, i) => i + 1).map((x) => {
    const today = new Date();
    const day = new Date(
      today.getTime() - 1000 * 60 * 60 * 24 * (days_back - x + 1)
    );
    const isToday =
      Math.abs(today.getTime() - day.getTime()) < 1000 * 60 * 60 * 24;
    if (isToday)
      display_sheet
        .getRange(2, x + 2, 2, 1)
        .setBorder(
          true,
          true,
          true,
          true,
          false,
          false,
          '#ffff00',
          SpreadsheetApp.BorderStyle.SOLID_MEDIUM
        );
    display_sheet
      .getRange(2, x + 2)
      .setValue(
        `'${day.getDate().toString().padStart(2, '0')}.${(day.getMonth() + 1)
          .toString()
          .padStart(2, '0')}`
      );
    display_sheet
      .getRange(3, x + 2)
      .setValue(
        day.toLocaleDateString('ru', { weekday: 'short' }).toUpperCase()
      );
  });
}

function getSettings() {
  const settings_sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
  const days_back = settings_sheet?.getRange(2, 1).getValue();
  const days_fw = settings_sheet?.getRange(2, 3).getValue();
  return { days_back, days_fw };
}

function renderEvents(props: {
  day_offset: number;
  days_fw: number;
  y_offset: number;
  x_offset: number;
  calendar_events: GoogleAppsScript.Calendar.CalendarEvent[];
  display_sheet: GoogleAppsScript.Spreadsheet.Sheet;
}) {
  const { day_offset, x_offset, y_offset, calendar_events, display_sheet } =
    props;
  const today = new Date();
  today.setHours(0);
  today.setMinutes(0);
  today.setSeconds(0);

  const cells = calendar_events
    .map((e) =>
      getCellsByTimeRange({
        day_offset,
        start: e.getStartTime() as Date,
        end: e.getEndTime() as Date,
        today,
        x_offset,
        y_offset,
      })
    )
    .flat();
  console.log(cells);
  cells.map(({ x, y }) =>
    display_sheet.getRange(y + 1, x + 1).setBackground(ACTIVE_COLOR)
  );
}

function getCalendarEvents(days_back: number, days_fw: number) {
  const calendar_id = PropertiesService.getScriptProperties().getProperty(
    CALENDAR_ID_PROP_NAME
  );
  if (!calendar_id) throw Error('CALENDAR_ID property not allowed');
  const calendar = CalendarApp.getCalendarById(calendar_id);
  if (!calendar) throw Error('Calendar not found');
  const start_time = new Date();
  start_time.setHours(0);
  start_time.setMinutes(0);
  start_time.setSeconds(0);
  start_time.setMilliseconds(0);

  start_time.setDate(start_time.getDate() - days_back);
  const end_time = new Date();
  end_time.setDate(end_time.getDate() + days_fw);
  const calendar_events = calendar?.getEvents(start_time, end_time) || [];
  return calendar_events;
}

function clearScreen(
  display_sheet: GoogleAppsScript.Spreadsheet.Sheet,
  days: number
) {
  const current_cnt = display_sheet.getMaxColumns();
  if (current_cnt > days + 2) {
    display_sheet.deleteColumns(days + 3, current_cnt - (days + 2));
  } else if (current_cnt < days + 2) {
    display_sheet.insertColumnsAfter(current_cnt, days + 2 - current_cnt);
    display_sheet
      .getRange(7, 3, 2, 1)
      .copyFormatToRange(
        display_sheet.getSheetId(),
        current_cnt + 1,
        display_sheet.getMaxColumns(),
        5,
        35
      );
  }

  const columns_cnt = display_sheet.getMaxColumns() - 2;
  display_sheet.getRange(4, 3, 32, columns_cnt).setBackground(INACTIVE_COLOR);
  display_sheet
    .getRange(2, 3, 3, columns_cnt)
    .setValue('')
    .setBorder(false, false, false, false, false, false);
}
