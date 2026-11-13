/** Heutiges Datum als Excel-Seriennummer (die Zelle behält ihr Datumsformat). */
export function excelToday(now = new Date()): number {
  return Math.round((Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) - Date.UTC(1899, 11, 30)) / 86400000);
}
