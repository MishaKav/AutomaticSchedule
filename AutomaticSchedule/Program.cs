using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using EAGetMail;
using Newtonsoft.Json;

namespace AutomaticSchedule
{
    public class Program
    {
        #region Properties

        private static readonly string dataPath = $@"{Directory.GetCurrentDirectory()}\Data\";
        private static readonly string currentFolder = $@"{dataPath}{DateTime.Now.ToString("dd-MM-yyyy")}\";

        // when we can use a gym
        private static readonly Dictionary<DayOfWeek, int> PossibleWorkShifts =
            new Dictionary<DayOfWeek, int>
            {
                { DayOfWeek.Sunday, 23 },
                { DayOfWeek.Monday, 23 },
                { DayOfWeek.Tuesday, 23 },
                { DayOfWeek.Wednesday, 23 },
                { DayOfWeek.Thursday, 23 },
                { DayOfWeek.Friday, 15 },
                { DayOfWeek.Saturday, 7 }
            }; 

        #endregion Properties

        public static void Main(string[] args)
        {
            var globalWatch = Stopwatch.StartNew();

            if (AppSettings.IsLocalWork)
            {
                //Utils.SendMailNotification($"<a href='{GoogleApi.GetGoogleCalendarEvent("Gym", DateTime.Now, null, "Gym")}' target='_blank'>Test Link</a><br/> some html");
                //var workSchedule = LoadLastResult();
                //SendEmailNotification(workSchedule);
                //AddWorkScheduleToCalendar(workSchedule);
            }
            else
            {
                if (SaveExcelFromGmail())
                {
                    var scheduleFiles = Directory.GetFiles(currentFolder, $"*.{AppSettings.WantedExtention}").Select(s => new FileInfo(s)).ToList();
                    if (scheduleFiles.IsAny())
                    {
                        var file = scheduleFiles.OrderBy(f => f.LastWriteTime).First();
                        Utils.RunProgressBar("Scaning Excel", 620);
                        var watch = Stopwatch.StartNew();
                        var workSchedule = ExcelApi.GetWorkHours(file.FullName, AppSettings.WantedWorker);
                        watch.Stop();
                        Utils.ProgressBar.Finish();

                        if (workSchedule.IsNotEmptyObject())
                        {
                            Utils.WriteStatus($"Success to analize Work Schedule | {watch.ToDefaultFormat()}");

                            AddWorkScheduleToCalendar(workSchedule);
                        }
                        else
                        {
                            Utils.SendMailNotification("Cann't find worker: " + AppSettings.WantedWorker);
                        }
                        SendEmailNotification(workSchedule);
                    }
                    else
                    {
                        Utils.SendMailNotification("Gmail scan return true, but no file founded");
                    }
                }
            }

            globalWatch.Stop();
            Utils.WriteStatus("Work Done | " + globalWatch.ToDefaultFormat());
        }

        // return all pair of trainigs
        private static List<Tuple<Reminder, Reminder>> GetTraningsPairs(WorkSchedule workSchedule)
        {
            if (workSchedule != null && workSchedule.Reminders.IsAny())
            {
                // get all days, that we can use a gym
                var relevantTranings = workSchedule.Reminders.FindAll(r => PossibleWorkShifts.Any(ws => ws.Key.Equals(r.Start.DayOfWeek) && ws.Value <= r.Start.Hour));

                if (relevantTranings.IsAny())
                {
                    var actualTranings = new List<Tuple<Reminder, Reminder>>();
                    foreach (var t1 in relevantTranings)
                    {
                        foreach (var t2 in relevantTranings)
                        {
                            if ((t2.Start - t1.Start).TotalHours > 24)
                            {
                                actualTranings.Add(new Tuple<Reminder, Reminder>(t1, t2));
                            }
                        }
                    }

                    return actualTranings;
                }
            }

            return null;
        }

        // add all reminders from schedule to calendar
        private static void AddWorkScheduleToCalendar(WorkSchedule workSchedule)
        {
            Utils.RunProgressBar("Connect to Google Calendar", 7);
            var coonected = GoogleApi.ConnectCalendar();
            Utils.ProgressBar.Finish();

            if (coonected)
            {
                Utils.RunProgressBar($"Start adding {workSchedule.Reminders.Count} event to Google Calendar", 10);
                var watch = Stopwatch.StartNew();
                foreach (var reminder in workSchedule.Reminders)
                {
                    var eventAdded = GoogleApi.AddEvent(reminder);
                    Utils.WriteStatus($"{reminder} is {(eventAdded ? "added" : "NOT ADDED")}");
                }
                watch.Stop();
                Utils.ProgressBar.Finish();

                Utils.WriteStatus($"Success add {workSchedule.Reminders.Count} events | {watch.ToDefaultFormat()}");

                var jsonSchedule = JsonConvert.SerializeObject(workSchedule, Formatting.Indented);
                File.WriteAllText($"{currentFolder}data.json", jsonSchedule);
            }
        }

        // send mail notification about added reminders + suggest days of trainigs
        private static void SendEmailNotification(WorkSchedule workSchedule)
        {
            if (workSchedule.IsNotEmptyObject() && workSchedule.Reminders.IsAny())
            {
                var list = workSchedule.Reminders.Select(r => r.ToDisplayFormat());
                var mailMessage = string.Join("<br/>", list) + "<br/><br/>";

                if (AppSettings.SuggestTrainigs)
                {
                    mailMessage += "<u>Suggested Trainings:</u><br/>";
                    var traningsPairs = GetTraningsPairs(workSchedule);

                    // only traning pairs
                    if (traningsPairs.IsAny())
                    {
                        var listPairs =
                            traningsPairs.Select(
                                t =>
                                    $"#{traningsPairs.IndexOf(t) + 1} <a href='{GoogleApi.GetGoogleCalendarEvent("Gym", t.Item1.Start.AddMinutes(30), t.Item1.Start.AddMinutes(180), "#Gym#")}' target='_blank'>{t.Item1.Start.ToString("dddd")}</a> " +
                                    $"({t.Item1.Start.ToDefaultDateFormat()})<br/>" +
                                    $"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='{GoogleApi.GetGoogleCalendarEvent("Gym", t.Item2.Start.AddMinutes(30), t.Item2.Start.AddMinutes(180), "#Gym#")}' target='_blank'>{t.Item2.Start.ToString("dddd")}</a> " +
                                    $"({t.Item2.Start.ToDefaultDateFormat()})<br/>").ToList();
                        mailMessage += string.Join("<br/>", listPairs) + "<br/><br/>";
                    }
                    else
                    {
                        // get all days, that we can use a gym
                        var relevantTranings = workSchedule.Reminders.FindAll(r => PossibleWorkShifts.Any(ws => ws.Key.Equals(r.Start.DayOfWeek) && ws.Value <= r.Start.Hour));

                        var listPairs =
                            relevantTranings.Select(
                                t =>
                                    $"#{relevantTranings.IndexOf(t) + 1} <a href='{GoogleApi.GetGoogleCalendarEvent("Gym", t.Start.AddMinutes(30), t.Start.AddMinutes(180), "#Gym#")}' target='_blank'>{t.Start.ToString("dddd")}</a> " +
                                    $"({t.Start.ToDefaultDateFormat()})<br/>").ToList();
                        mailMessage += string.Join("<br/>", listPairs) + "<br/><br/>";
                    }
                }
                
                mailMessage += "Automatic Schedule by Misha Kav :)";
                Utils.WriteStatus(mailMessage);

                if (AppSettings.SendMailNotification)
                {
                    Utils.SendMailNotification(mailMessage);
                }
            }
        }

        // need to work from local. Just load the last result from local json file
        private static WorkSchedule LoadLastResult()
        {
            var jsonFiles = Directory.GetFiles(dataPath, "*.json", SearchOption.AllDirectories).Select(s => new FileInfo(s)).ToList();

            if (jsonFiles.IsAny())
            {
                var file = jsonFiles.OrderByDescending(f => f.LastWriteTime).First();
                var jsonData = File.ReadAllText(file.FullName);
                if (jsonData.IsNotNullOrEmpty())
                {
                    var workSchedule = JsonConvert.DeserializeObject<WorkSchedule>(jsonData);
                    return workSchedule;
                }

            }
            return null;
        }

        // scan and save Gmail from attachment with wanted format
        private static bool SaveExcelFromGmail()
        {
            Utils.RunProgressBar("Connecting to Gmail", 150);
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
                    if (oMail.ReceivedDate > DateTime.Now.AddMonths(-1))
                    {
                        Utils.WriteStatus($"Date: {oMail.ReceivedDate.ToDefaultDateTimeFormat()} | From: {oMail.From} | Subject: {oMail.Subject.Replace("(Trial Version)", string.Empty)}");

                        if (oMail.Attachments.IsAny() && oMail.Attachments.Any(a => a.Name.ContainsIgnoreCase(AppSettings.WantedExtention)))
                        {
                            if (!Directory.Exists(currentFolder))
                            {
                                Directory.CreateDirectory(currentFolder);
                            }

                            //var fileName = $@"{currentFolder}/{oMail.From.Address}_{oMail.Subject.Replace(":", string.Empty).Replace("(Trial Version)", string.Empty)}.eml";
                            oMail.Attachments[0].SaveAs($@"{currentFolder}\{oMail.Attachments[0].Name}", true);
                            //oMail.SaveAs(fileName, true);
                            findAttachment = true;
                            break;
                        }
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