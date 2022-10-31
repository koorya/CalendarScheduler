function getCellsByTimeRange({
  today,
  day_offset,
  start,
  end,

  x_offset,
  y_offset,
}: {
  today: Date;
  day_offset: number;
  start: Date;
  end: Date;
  x_offset: number;
  y_offset: number;
}) {
  const time_offset = 9;
  if (end.getDate() !== start.getDate()) return [];

  const day =
    day_offset +
    Math.floor((start.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
  if (day < 0) return [];
  const y_start =
    y_offset -
    2 * time_offset +
    2 * start.getHours() +
    Math.floor(start.getMinutes() / 30);
  const y_end =
    y_offset -
    2 * time_offset +
    2 * end.getHours() +
    Math.ceil(end.getMinutes() / 30);
  const cell_cnt = y_end - y_start;
  return Array.from({ length: cell_cnt }, (x, i) => ({
    x: day + x_offset,
    y: i + y_start,
  })).filter(({ y }) => y > y_offset);
}
