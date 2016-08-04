using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using Redmine.Net.Api;
using Redmine.Net.Api.Types;
using System.Drawing;
using System.Globalization;


namespace RedmineTimeConsumer
{
    class Program
    {
        static double _underTimeLimit = -0.5;
        static double _overTimeLimit = 4;
        static double _normalTime = 8;
        static string redisImageName = "redis.jpg";

        static List<string> loginList = new List<string>();

        static Dictionary<string, string> nameDictionary = new Dictionary<string, string>();

        static Dictionary<string, DateTime> userHolidays = new Dictionary<string, DateTime>();
        static Dictionary<string, string> userTypes = new Dictionary<string, string>();
        static Dictionary<string, int> userRedises = new Dictionary<string, int>();

        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("Usage: RedMineTimeConsumer <Дата yyyy-MM-dd>");
                Environment.Exit(-1);
            }

            //var login = args[0];
            string dateStr = args.Length < 1 ? DateTime.Now.ToString("yyyy-MM-dd") : args[0];
            DateTime date = DateTime.Now;
            DateTime.TryParse(dateStr, out date);

            string host = "http://redmine.ugsk.ru";
            string apiKey = "1befa3b66839aa5dba8668ba7d1e4c873147b4a8";

            var manager = new RedmineManager(host, apiKey);

            var pathTime = string.Format("{0}-time.html", dateStr);
            var pathReport = string.Format("{0}-report.html", dateStr);

            // Читаем файл с маппингом имён
            string cfgName = "NameMapping.cfg";
            try
            {
                string[] nameMappingLines = System.IO.File.ReadAllLines(cfgName, Encoding.UTF8);
                foreach (var nml in nameMappingLines)
                {
                    string[] line = nml.Split(new string[] { ";" }, StringSplitOptions.None);
                    nameDictionary.Add(line[0], line[1]);
                }
            }
            catch (Exception ex)
            {
                System.Console.WriteLine(string.Format("Ошибка разбора файла {0}: {1}", cfgName, ex.Message));
            }

            // Читаем файл с отпусками
            cfgName = "RedMineTimeConsumer.cfg";
            try
            {
                string[] userConfigLines = System.IO.File.ReadAllLines(cfgName, Encoding.UTF8);
                foreach (var ucl in userConfigLines)
                {
                    string[] line = ucl.Split(new string[] { ";" }, StringSplitOptions.None);
                    loginList.Add(line[0]);
                    userTypes.Add(line[0], line[1]);

                    DateTime endHolidaysDate = DateTime.MinValue;
                    DateTime.TryParse(line[2], out endHolidaysDate);
                    endHolidaysDate = endHolidaysDate.AddSeconds(24 * 60 * 60 - 1); // Конец дня
                    int redisCount = 0;
                    int.TryParse(line[3], out redisCount);
                    endHolidaysDate = endHolidaysDate.AddSeconds(24 * 60 * 60 - 1); // Конец дня
                    userHolidays.Add(line[0], endHolidaysDate);
                    userRedises.Add(line[0], redisCount);
                }
            }
            catch (Exception ex)
            {
                System.Console.WriteLine(string.Format("Ошибка разбора файла {0}: {1}", cfgName, ex.Message));
            }

            //var parameters = new NameValueCollection { { "limit", "1000" }, { "offset", "0" } }; // Почему-то возвращало только 100 записей 
            var parameters = new NameValueCollection { };

            using (var swReport = new StreamWriter(pathReport, false, Encoding.Default))
            {
                swReport.WriteLine("<style>span{color:gray; display:inline; font-size:11pt}</style>");
                swReport.WriteLine("<div style=\"font-size:9pt; display:inline\"><div style=\"color:gray; display:inline\">Данные Redmine</div>&nbsp;Отчёт сотрудника</div>");

                using (var swTime = new StreamWriter(pathTime, false, Encoding.Default))
                {
                    //swTime.WriteLine("<style>span{color:gray; display:inline; font-size:11pt}</style>");
                    swTime.WriteLine(string.Format("<br/>Учтённые трудозатраты за {0:dd.MM.yyyy}<br/>по состоянию на {1:dd.MM.yyyy HH:mm:ss}", date, DateTime.Now));
                    swTime.WriteLine(@"<br/><br/><table border=""1"" cellpadding=""2"" style=""border-collapse: collapse; border: 1px solid black;"">");
                    swTime.WriteLine(@"<style type=""text/css"">TD {text-align: center;} TR.gray {color: Gray;} TD.redis {text-align: left; width: 300;}</style>");

                    //swTime.WriteLine("<tr><th>ФИО</th><th>Трудозатраты, ч</th><th>Отклонение, ч</th></tr>");

                    var allUsers = manager.GetTotalObjectList<User>(parameters);

                    foreach (var login in loginList)
                    {
                        // Проверяем на отпуск
                        DateTime endHolidaysDate = userHolidays[login];
                        bool isHoliday = date.CompareTo(endHolidaysDate.Date) < 0;

                        // Читаем количество редисок
                        int redisCount = userRedises[login];

                        //var users = manager.GetObjectList<User>(parameters).Where(e => e.LastName.Contains(employeeName)
                        //                                                           && !e.FirstName.Contains("External")
                        //                                                            ).OrderBy(e => e.Id);

                        // Определим ID пользователя

                        //var users = manager.GetTotalObjectList<User>(parameters).Where(e => e.Login.Equals(login, StringComparison.InvariantCultureIgnoreCase)
                        //                                    ).OrderBy(e => e.Id);

                        //int userID = 0;
                        //string userName = string.Empty;

                        //int c = users.Count();

                        //foreach (var u in users)
                        //{
                        //    userID = u.Id;
                        //    userName = string.Format("{0} {1}", u.LastName, u.FirstName);
                        //    break; // instead of top 1 :)
                        //}
                        

                        int userID = 0;
                        string userName = string.Empty;
                        foreach (var u in allUsers)
                        {
                            if (login.Equals(u.Login, StringComparison.InvariantCultureIgnoreCase))
                            {
                                userID = u.Id;
                                userName = string.Format("{0} {1}", u.LastName, u.FirstName);
                                break;
                            }
                        }

                        // Time Entries

                        parameters = new NameValueCollection { { "spent_on", dateStr }
                                                  ,{ "limit", "1000" }
                                                  ,{ "offset", "0" }
                                                  ,{ "user_id", userID.ToString() }
                                                 };
                        DateTime d = DateTime.Now;
                        DateTime.TryParse(dateStr, out d);

                        var timeEntries = manager.GetObjectList<TimeEntry>(parameters).Where(e => e.SpentOn.Equals(d))
                                                            .OrderBy(e => e.Id);

                        int count = timeEntries.Count();

                        int i = 0;

                        var strBuilder = new StringBuilder();

                        //bool isWrittenEmployeeName = false;

                        //var path = Path.GetTempFileName() + ".txt";
                        StringBuilder sbOut = new StringBuilder();
                        double hoursSum = 0;
                        //string name = string.Empty;

                        //sw.WriteLine("--------------");
                        sbOut.AppendLine("<br/><br/><b>" + nameDictionary[userName] + "</b>");

                        if (isHoliday)
                        {
                            sbOut.AppendLine(string.Format("<br/>Отпуск до {0:dd.MM}", endHolidaysDate));
                        }
                        else
                        {
                            foreach (var timeEntry in timeEntries) //
                            {

                                //if (!isWrittenEmployeeName)
                                //{
                                //    //sw.Write(timeEntry.User.Name);
                                //    //name = timeEntry.User.Name;
                                //    sbOut.AppendLine(userName);
                                //    isWrittenEmployeeName = true;
                                //}
                                strBuilder.Clear();
                                strBuilder.AppendFormat("<br/><span>{0}. {1} ({2}): ", ++i, timeEntry.Project.Name, timeEntry.Activity.Name, timeEntry.SpentOn);
                                //Issue issue=null;
                                if (timeEntry.Issue != null)
                                {
                                    var issue = manager.GetObject<Issue>(timeEntry.Issue.Id.ToString(), new NameValueCollection());
                                    strBuilder.AppendFormat("{0} <a href=\"http://redmine.ugsk.ru/issues/{1}\">RM#{1}</a>", issue.Subject, issue.Id);
                                    if (!string.IsNullOrEmpty(timeEntry.Comments))
                                        strBuilder.Append(" - ");
                                }

                                strBuilder.AppendFormat("</span>{0} <span>({1} ч.);</span>", timeEntry.Comments, timeEntry.Hours);
                                hoursSum += (double)timeEntry.Hours;
                                //sw.WriteLine(strBuilder.ToString());
                                sbOut.AppendLine(strBuilder.ToString());
                            }
                        }
                        swReport.WriteLine(sbOut.ToString());


                        // Формируем код с редисками
                        StringBuilder redisStringBuilder = new StringBuilder();
                        for (int r = 0; r < redisCount; r++)
                        {
                            redisStringBuilder.Append(string.Format(@"<img src=""{0}""/>", redisImageName));
                        }
                        string redisHtml = redisStringBuilder.ToString();

                        if (isHoliday)
                        {
                            swTime.WriteLine(@"<tr class=""gray""><td>{0}</td><td colspan=""2"">отпуск до {1:dd.MM}</td><td class=""redis"" bgcolor=""White"">{2}</td></tr>", userName, endHolidaysDate, redisHtml);
                        }
                        else
                        {
                            // Определяем индивидуальную норму в заданный день для сотрудника
                            double normalTime = _normalTime;
                            string type = userTypes[login];
                            if ("8".Equals(type, StringComparison.InvariantCultureIgnoreCase))
                            {
                                normalTime = 8;
                            }
                            else if (("F".Equals(type, StringComparison.InvariantCultureIgnoreCase)
                                && date.DayOfWeek == DayOfWeek.Friday)
                                || ("4".Equals(type, StringComparison.InvariantCultureIgnoreCase)))
                            {
                                normalTime = 4;
                            }

                            string overTime = string.Empty;
                            double overTimeValue = hoursSum - normalTime;
                            if (overTimeValue < 0)
                            {
                                overTime = string.Format("{0:0.00}", overTimeValue);
                            }
                            else if (overTimeValue > 0)
                            {
                                overTime = string.Format("+{0:0.00}", overTimeValue);
                            }
                            else
                            {
                                overTime = string.Empty;
                            }

                            if (overTimeValue >= _underTimeLimit)
                            {
                                if (overTimeValue <= 0)
                                {
                                    swTime.WriteLine(@"<tr><td>{0}</td><td colspan=""2"">{1}</td><td class=""redis"">{2}</td></tr>", userName, hoursSum, redisHtml);
                                }
                                else if (overTimeValue > 0)
                                {
                                    swTime.WriteLine(@"<tr><td>{0}</td><td>{1}</td><td bgcolor=""{3}"">{2}</td><td class=""redis"">{4}</td></tr>", userName, hoursSum, overTime, GetHtmlColor(Color.Green, GetAlphaPercent(overTimeValue, _overTimeLimit)), redisHtml);
                                }
                            }
                            else
                            {
                                if (hoursSum == 0)
                                {
                                    swTime.WriteLine(@"<tr><td bgcolor=""red"">{0}</td><td colspan=""2"" bgcolor=""red"">-</td><td class=""redis"">{1}</td></tr>", userName, redisHtml);
                                }
                                else
                                {
                                    swTime.WriteLine(@"<tr><td bgcolor=""red"">{0}</td><td bgcolor=""{3}"">{1}</td><td bgcolor=""{3}"">{2}</td><td class=""redis"">{4}</td></tr>", userName, hoursSum, overTime, GetHtmlColor(Color.Red, GetAlphaPercent(-overTimeValue, normalTime)), redisHtml);
                                }
                            }

                        }
                        swTime.Flush();
                        swReport.Flush();
                    }
                    swTime.WriteLine("</table>");
                }
                //var p = Process.Start("notepad.exe", path);

                //if (p != null)
                //{
                //    p.WaitForExit();
                //}

                //File.Delete(path);

                //Create a issue.
                //var newIssue = new Issue { Subject = "test", Project = new IdentifiableName { Id = 1 } };
                //manager.CreateObject(newIssue);
            }
        }

        private static Color Interpolate(Color color1, Color color2, double fraction)
        {
            Color c = ColorInterpolator.InterpolateBetween(color1, color2, fraction);
            return c;
        }

        private static String GetHtmlColor(Color baseColor, double fraction)
        {
            Color color = Interpolate(Color.White, baseColor, fraction);
            string res = System.Drawing.ColorTranslator.ToHtml(color);
            return res;
        }

        private static double GetAlphaPercent(double overTime, double timeLimit)
        {
            double r = overTime / timeLimit;
            if (r > 1) r = 1;
            return r;
        }

    }
}
