using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using EAGetMail;

namespace AutomaticSchedule
{
    public class Program
    {
        private static readonly string dataPath = $@"{Directory.GetCurrentDirectory()}\Data\";
        private static readonly string currentFolder = $@"{dataPath}{DateTime.Now.ToString("dd-MM-yyyy")}/";
        private const string wanterExtention = "xlsx";
        private const string wanterWorker = "קוסטיה";

        public static void Main(string[] args)
        {

            Utils.RunProgressBar("Connecting to Gmail", 150);
            if (SaveExcelFromGmail())
            {
                var scheduleFiles = Directory.GetFiles(currentFolder, $"*.{wanterExtention}").Select(s => new FileInfo(s)).ToList();
                if (scheduleFiles.IsAny())
                {
                    var file = scheduleFiles.OrderBy(f => f.LastWriteTime).First();
                    Utils.RunProgressBar("Scaning Excel", 620);
                    var watch = Stopwatch.StartNew();
                    var _workSchedule = ExcelApi.GetWorkHours(file.FullName, wanterWorker);
                    watch.Stop();
                    Utils.ProgressBar.Finish();

                    if (_workSchedule.IsNotEmptyObject())
                    {
                        Utils.WriteStatus($"Success to analize Work Schedule | {watch.ToDefaultFormat()}");

                        Utils.RunProgressBar("Connect to Google Calendar", 7);
                        var coonected = GoogleApi.ConnectCalendar();
                        Utils.ProgressBar.Finish();

                        if (coonected)
                        {
                            Utils.RunProgressBar($"Start adding {_workSchedule.Reminders.Count} event to Google Calendar", 10);
                            var watch1 = Stopwatch.StartNew();
                            foreach (var reminder in _workSchedule.Reminders)
                            {
                                var eventAdded = GoogleApi.AddEvent(reminder);
                                Utils.WriteStatus($"{reminder} is {(eventAdded ? "added" : "NOT ADDED")}");
                            }
                            watch1.Stop();
                            Utils.ProgressBar.Finish();

                            Utils.WriteStatus($"Success add {_workSchedule.Reminders.Count} events | {watch1.ToDefaultFormat()}");

                            var list = GetPrintedSchedule(_workSchedule);
                            list.ForEach(Utils.WriteStatus);
                            Utils.SendMailNotification(string.Join("\n", list) + "\n\nAutomatic Schedule by Misha Kav :)");
                        }
                    }
                    else
                    {
                        Utils.SendMailNotification("Cann't find worker: " + wanterWorker);
                    }
                }
                else
                {
                    Utils.SendMailNotification("Gmail scan return true, but no file founded");
                }
            }
        }

        private static List<string> GetPrintedSchedule(WorkSchedule workSchedule)
        {
            var list = new List<string>();

            foreach (var reminder in workSchedule.Reminders)
            {
                var text = $"{reminder.Start.ToDefaultDateFormat()}  {reminder.Start.ToString("ddd"), -8}: {reminder.Start.ToDefaultTimeFormat()} - {reminder.End.ToDefaultTimeFormat()}    [{(reminder.End - reminder.Start).TotalHours,-2} h]";
                list.Add(text);
            }

            return list;
        }

        private static bool SaveExcelFromGmail()
        {
            if (!Directory.Exists(currentFolder))
            {
                Directory.CreateDirectory(currentFolder);
            }

            var findAttachment = false;
            var oServer = new MailServer("imap.gmail.com", "developer.newconcept@gmail.com", "robot555", ServerProtocol.Imap4);
            var oClient = new MailClient("TryIt");
            oServer.SSLConnection = true;
            oServer.Port = 993;

            try
            {
                oClient.Connect(oServer);
                Utils.ProgressBar.Finish();
                var infos = oClient.GetMailInfos();
                foreach (var info in infos.Reverse())
                {
                    //Console.WriteLine($"Index: {info.Index}; Size: {info.Size}; UIDL: {info.UIDL}");
                    var oMail = oClient.GetMail(info);
                    Utils.WriteStatus($"Date: {oMail.ReceivedDate.ToDefaultDateTimeFormat()} | From: {oMail.From} | Subject: {oMail.Subject.Replace("(Trial Version)", string.Empty)}");

                    if (oMail.Attachments.IsAny() && oMail.Attachments.Any(a => a.Name.ContainsIgnoreCase(wanterExtention)))
                    {
                        //var fileName = $@"{currentFolder}/{oMail.From.Address}_{oMail.Subject.Replace(":", string.Empty).Replace("(Trial Version)", string.Empty)}.eml";
                        oMail.Attachments[0].SaveAs($@"{currentFolder}\{oMail.Attachments[0].Name}", true);
                        //oMail.SaveAs(fileName, true);
                        findAttachment = true;
                        break;
                    }
                }

                oClient.Quit();
                return findAttachment;
            }
            catch (Exception ex)
            {
                Utils.WriteErrorLog(ex, "Scan Gmail");
            }

            return false;
        }
    }
}