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

namespace RedmineTimeConsumer
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("Не указан ты");
                Environment.Exit(-1);
            }

            var employeeName = args[0];
            string date = args.Length < 2 ? DateTime.Now.ToString("yyyy-MM-dd") : args[1];

            string host = "http://redmine.ugsk.ru";
            //string apiKey = "cf312b768b25dcd8d6de335762f359d891544a28";
            string apiKey = "1befa3b66839aa5dba8668ba7d1e4c873147b4a8";

            var manager = new RedmineManager(host, apiKey);

            // Определим ID пользователя
            var parameters = new NameValueCollection { { "limit", "1000" }
                                                      ,{ "offset", "0" }
                                                    };
            var users = manager.GetObjectList<User>(parameters).Where(e => e.LastName.Contains(employeeName)
                                                                       && !e.FirstName.Contains("External")
                                                                        ).OrderBy(e => e.Id);
            int userID = 0;
            foreach (var u in users)
            {
                userID = u.Id;
            }

            // Time Entries

            parameters = new NameValueCollection { { "spent_on", date }
                                                  ,{ "limit", "1000" }
                                                  ,{ "offset", "0" }
                                                  ,{ "user_id", userID.ToString() }
                                                 };
            DateTime d = DateTime.Now;
            DateTime.TryParse(date, out d);

            var timeEntries = manager.GetObjectList<TimeEntry>(parameters).Where(e => e.SpentOn.Equals(d))
                                                .OrderBy(e => e.CreatedOn);

            int count = timeEntries.Count();

            int i = 0;

            var strBuilder = new StringBuilder();

            bool isWrittenEmployeeName = false;

            //var path = Path.GetTempFileName() + ".txt";
            var pathTime = string.Format("{0}-time.txt", date);
            var pathReport = string.Format("{0}-report.txt", date);
            StringBuilder sbOut = new StringBuilder();
            double hoursSum = 0;
            string name = string.Empty;

            using (var swReport = new StreamWriter(pathReport, true))
            {
                using (var swTime = new StreamWriter(pathTime, true))
                {
                    //sw.WriteLine("--------------");
                    sbOut.AppendLine("--------------");

                    foreach (var timeEntry in timeEntries) //
                    {

                        if (!isWrittenEmployeeName)
                        {
                            //sw.Write(timeEntry.User.Name);
                            name = timeEntry.User.Name;
                            sbOut.AppendLine(name);
                            isWrittenEmployeeName = true;
                        }
                        strBuilder.Clear();
                        strBuilder.AppendFormat("{0}. {1} ({2}): ", ++i, timeEntry.Project.Name, timeEntry.Activity.Name, timeEntry.SpentOn);
                        //Issue issue=null;
                        if (timeEntry.Issue != null)
                        {
                            var issue = manager.GetObject<Issue>(timeEntry.Issue.Id.ToString(), new NameValueCollection());
                            strBuilder.AppendFormat("{0} (RM#{1})", issue.Subject, issue.Id);
                            if (!string.IsNullOrEmpty(timeEntry.Comments))
                                strBuilder.Append(" - ");
                        }

                        strBuilder.AppendFormat("{0} ({1} ч.);", timeEntry.Comments, timeEntry.Hours);
                        hoursSum += (double)timeEntry.Hours;
                        //sw.WriteLine(strBuilder.ToString());
                        sbOut.AppendLine(strBuilder.ToString());

                    }

                    string overTime = string.Empty;
                    if ((8 - hoursSum) > 0)
                    {
                        overTime = string.Format("[-{0:0.00}]", 8 - hoursSum);
                    }
                    else if ((8 - hoursSum) < 0)
                    {
                        overTime = string.Format("[+{0:0.00}]", hoursSum - 8);
                    }
                    else
                    {
                        overTime = string.Empty;
                    }

                    if (!string.IsNullOrEmpty(name))
                    {
                        swTime.WriteLine("{0,15} [{1:0.00}] {2}", employeeName, hoursSum, overTime);
                    }
                    else
                    {
                        swTime.WriteLine("{0,15} [----]", employeeName);

                    }
                    //System.Console.WriteLine(sbOut.ToString());
                    swReport.WriteLine(sbOut.ToString());

                }
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
}
