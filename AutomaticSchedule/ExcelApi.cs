using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutomaticSchedule
{
    public static class ExcelApi
    {
        private static Excel.Range _range; // current range of curent excel

        public static WorkSchedule GetWorkHours(string filePath, string workerName)
        {
            var xlApp = new Excel.Application();
            var xlWorkBook = xlApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            var xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item[1];

            _range = xlWorkSheet.UsedRange;

            var name = workerName.Trim();
            var counter = 0;

            var workSchedule = new WorkSchedule
            {
                Name = name,
                //StartDateTime = Convert.ToDateTime(GetCellValue(4, 2)),
                //EndDateTime = Convert.ToDateTime(GetCellValue(4, 14))
                StartDateTime = DateTime.ParseExact(GetCellValue(4, 2), "d.M.yyyy", CultureInfo.InvariantCulture),
                EndDateTime = DateTime.ParseExact(GetCellValue(22, 2), "d.M.yyyy", CultureInfo.InvariantCulture)
            };

            for (var i = 1; i <= _range.Rows.Count; i++)
            {
                for (var j = 1; j <= _range.Columns.Count; j++)
                {
                    var cell = GetCellValue(i, j);

                    #region Old Discount
                    //if (cell == name)
                    //{
                    //    counter++;
                    //    string[] start = null, end = null;
                    //    var cellTimes = GetCellValue(i - 1, j);
                    //    if (!string.IsNullOrEmpty(cellTimes) && cellTimes.Contains(":"))
                    //    {
                    //        start = cellTimes.Split(':');
                    //    }

                    //    cellTimes = GetCellValue(i - 1, j + 1);
                    //    if (!string.IsNullOrEmpty(cellTimes) && cellTimes.Contains(":"))
                    //    {
                    //        end = cellTimes.Split(':');
                    //    }

                    //    var date = Convert.ToDateTime(GetCellValue(4, j));
                    //    if (start != null && end != null)
                    //    {
                    //        var reminder = new Reminder
                    //        {
                    //            DayDesc = GetCellValue(3, j),
                    //            Start = date.AddHours(int.Parse(start[0])).AddMinutes(int.Parse(start[1])),
                    //            End = date.AddHours(int.Parse(end[0])).AddMinutes(int.Parse(end[1])),
                    //            JobName = GetCellValue(i - 1, 1),
                    //        };

                    //        // check if end time in next day
                    //        if ((reminder.End - reminder.Start).TotalHours < 0)
                    //        {
                    //            reminder.End = reminder.End.AddDays(1);
                    //        }

                    //        workSchedule.Reminders.Add(reminder);
                    //    }
                    //} 
                    #endregion

                    #region CheckPoint
                    if (cell == name)
                    {
                        counter++;
                        int[] start = null, end = null;
                        var cellTimes = GetCellValue(i, 3);
                        if (!string.IsNullOrEmpty(cellTimes))
                        {
                            cellTimes = cellTimes.Trim();

                            if (cellTimes.Equals("בוקר"))
                            {
                                start = new[] { 07, 00 };
                                end = new[] { 15, 00 };
                            }

                            if (cellTimes.Equals("צהריים"))
                            {
                                start = new[] { 15, 00 };
                                end = new[] { 23, 00 };
                            }

                            if (cellTimes.Equals("לילה"))
                            {
                                start = new[] { 23, 00 };
                                end = new[] { 07, 00 };
                            }
                        }
                        var date = TryGetDate(i, 2);
                        if (start.IsAny() && end.IsAny())
                        {
                            var reminder = new Reminder
                            {
                                DayDesc = TryGetDayDesc(i, 2),
                                Start = date.AddHours(start[0]).AddMinutes(start[1]),
                                End = date.AddHours(end[0]).AddMinutes(end[1]),
                                JobName = GetCellValue(2, j)
                            };

                            // check if end time in next day
                            if ((reminder.End - reminder.Start).TotalHours < 0)
                            {
                                reminder.End = reminder.End.AddDays(1);
                            }

                            workSchedule.Reminders.Add(reminder);
                        }
                    }
                    #endregion
                }
            }

            workSchedule.Reminders = workSchedule.Reminders.OrderBy(r => r.Start).ToList();

            //MessageBox.Show(text);

            xlWorkBook.Close(true);
            xlApp.Quit();

            ReleaseObject(xlWorkSheet);
            ReleaseObject(xlWorkBook);
            ReleaseObject(xlApp);

            if (counter > 0)
            {
                var json = JsonConvert.SerializeObject(workSchedule, Formatting.Indented);
                var fileName =
                    $@"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\{
                        workSchedule.StartDateTime.ToString("dd.MM.yy")}-{workSchedule.EndDateTime.ToString("dd.MM.yy")
                        }_{name}.txt";

                if (Utils.IsMishaPc())
                {
                    File.WriteAllText(fileName, json);
                }

                return workSchedule;
            }
            else
            {
                return null;
            }
        }

        private static string TryGetDayDesc(int i, int j)
        {
            if (i >= 3 && i <= 5)
            {
                i = 4;
            }
            else if (i >= 6 && i <= 8)
            {
                i = 7;
            }
            else if (i >= 9 && i <= 11)
            {
                i = 10;
            }
            else if (i >= 12 && i <= 14)
            {
                i = 13;
            }
            else if (i >= 15 && i <= 17)
            {
                i = 16;
            }
            else if (i >= 18 && i <= 20)
            {
                i = 19;
            }
            else if (i >= 21 && i <= 23)
            {
                i = 22;
            }

            if (GetCellValue(i, j) != null)
            {
                return GetCellValue(i, j);
            }

            return string.Empty;
        }

        private static DateTime TryGetDate(int i, int j)
        {
            if (i >= 3 && i <= 5)
            {
                i = 4;
            }
            else if (i >= 6 && i <= 8)
            {
                i = 7;
            }
            else if (i >= 9 && i <= 11)
            {
                i = 10;
            }
            else if (i >= 12 && i <= 14)
            {
                i = 13;
            }
            else if (i >= 15 && i <= 17)
            {
                i = 16;
            }
            else if (i >= 18 && i <= 20)
            {
                i = 19;
            }
            else if (i >= 21 && i <= 23)
            {
                i = 22;
            }

            if (GetCellValue(i, j) != null)
            {
                return DateTime.ParseExact(GetCellValue(i, j), "d.M.yyyy", CultureInfo.InvariantCulture);
            }

            return DateTime.Now;
        }

        private static string GetCellValue(int i, int j)
        {
            if (i > 0 && j > 0)
            {
                var cell = _range.Cells[i, j] as Excel.Range;
                if (!string.IsNullOrWhiteSpace(cell?.Text))
                {
                    return cell.Text.Trim();
                }
            }

            return null;
        }

        private static void ReleaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
            }
            catch (Exception ex)
            {
                Utils.WriteErrorLog(ex, "Unable to release the Object");
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
