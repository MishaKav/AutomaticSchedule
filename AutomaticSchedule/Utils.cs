using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using NLog;

namespace AutomaticSchedule
{
    public static class Utils
    {
        #region NLOG
        public enum LoggerList
        {
            ErrorLog = 0,       // defaulft logger - see NGLOG.config
            MyCustomLog,
            StatusLog,
            Gmail
        }

        public static Logger GetLogger(LoggerList logName = LoggerList.ErrorLog)
        {
            return LogManager.GetLogger(logName.ToString());
        }

        public static void WriteErrorLog(string msg)
        {
            if (msg.IsNullOrEmpty())
            {
                LogManager.GetLogger(LoggerList.ErrorLog.ToString()).Debug("Empty Message");
            }

            LogManager.GetLogger(LoggerList.ErrorLog.ToString()).Debug(msg);
            LogManager.GetLogger(LoggerList.ErrorLog.ToString()).Debug(string.Empty);
        }

        public static void WriteErrorLog(Exception error, string additionalMsg = "")
        {
            var addMsg = string.Empty;

            if (error.InnerException != null)
                error = error.InnerException;

            if (additionalMsg.IsNotNullOrEmpty())
                addMsg = $@"Additional Information: {additionalMsg}";

            var msg =
                $@"=================== Exception ==================
                    Exception message:
                    {error.Message}
                    Exception stack trace:
                    {error.StackTrace}
                    {addMsg}
                    -----------------------------------------------";

            LogManager.GetLogger(LoggerList.ErrorLog.ToString()).Debug(msg);
            LogManager.GetLogger(LoggerList.ErrorLog.ToString()).Debug(string.Empty);
        }

        public static void WriteErrorLog(LoggerList logName = LoggerList.ErrorLog, string additionalMsg = "")
        {
            var logger = LogManager.GetLogger(logName.ToString());
            logger.Debug(additionalMsg);
            logger.Debug(string.Empty);
        }

        public static void WriteMyCustomLog(string body, string title = "Test")
        {
            var msg = $@"{body}";
            LogManager.GetLogger(LoggerList.MyCustomLog.ToString()).Debug(msg);
            msg = @"===========================================";
            LogManager.GetLogger(LoggerList.MyCustomLog.ToString()).Debug(msg);
        }

        public static void WriteStatus(string msg)
        {
            Console.WriteLine(msg);
            LogManager.GetLogger(LoggerList.StatusLog.ToString()).Info(msg);
        }

        /// <summary>
        /// Just send email notification by Gmail account from NLog.config
        /// </summary>
        /// <param name="msg">message to send</param>
        public static void SendMailNotification(string msg)
        {
            var fullMsg = $"Send email notification: {msg}";
            Console.WriteLine(fullMsg);
            WriteStatus(fullMsg);

            // send email by Gmail
            LogManager.GetLogger(LoggerList.Gmail.ToString()).Info(msg);
        }
        #endregion NLOG

        #region ProgressBar

        public static ProgressBar ProgressBar;

        public static void RunProgressBar(string title = "", int delay = 20)
        {
            Task.Factory.StartNew(() =>
            {
                Console.WriteLine(title);
                using (ProgressBar = new ProgressBar())
                {
                    for (var i = 0; i <= 100; i++)
                    {
                        var value = (double) i/100;

                        // other classes can interrupt this progress for finish
                        if (ProgressBar.CurrentProgress < 1)
                        {
                            ProgressBar.Report(value);
                            Thread.Sleep(delay);
                        }
                        else
                        {
                            ProgressBar.Report(1);
                            break;
                        }
                    }
                }
            });
        }

        #endregion ProgressBar

        #region Misc

        public static string GetWebRequest(string url)
        {
            var req = (HttpWebRequest)WebRequest.Create(url);
            req.Timeout = 650000;
            req.UserAgent = "Misha Kav";
            req.Method = "GET";

            HttpWebResponse rsp = null;
            try
            {
                rsp = (HttpWebResponse)req.GetResponse();
            }
            catch (WebException wException)
            {
                if (wException.Status == WebExceptionStatus.ProtocolError)
                    rsp = (HttpWebResponse)wException.Response;
            }

            var sres = "";
            if (rsp != null)
            {
                try
                {
                    var dataStream = rsp.GetResponseStream();
                    if (dataStream != null)
                    {
                        var sread = new StreamReader(dataStream);
                        sres = sread.ReadToEnd();
                    }
                }
                catch
                {
                    // ignored
                }
            }

            return sres;
        }

        public static bool IsMishaPc()
        {
            var arr = new[] { "misha-pc", "misha-dell" };
            return arr.Contains(Environment.MachineName.ToLower());
        }

        #endregion Misc
    }
}
