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
            string apiKey = "cf312b768b25dcd8d6de335762f359d891544a28";

            var manager = new RedmineManager(host, apiKey);

            var parameters = new NameValueCollection { { "spent_on", date }, { "limit", "1000" }, { "offset", "0" } };
            var timeEntries = manager.GetObjectList<TimeEntry>(parameters).Where(e => e.User.Name.Contains(employeeName)).OrderBy(e => e.CreatedOn);


            int i = 0;

            var strBuilder = new StringBuilder();

            bool isWrittenEmployeeName = false;


            var path = Path.GetTempFileName() + ".txt";
            using (var sw = new StreamWriter(path))
            {
                sw.WriteLine("--------------");

                foreach (var timeEntry in timeEntries) //
                {
                    if (!isWrittenEmployeeName)
                    {
                        sw.WriteLine(timeEntry.User.Name);
                        isWrittenEmployeeName = true;
                    }
                    strBuilder.Clear();
                    strBuilder.AppendFormat("{0}. {1} ({2}): ", ++i, timeEntry.Project.Name, timeEntry.Activity.Name);
                    //Issue issue=null;
                    if (timeEntry.Issue != null)
                    {
                        var issue = manager.GetObject<Issue>(timeEntry.Issue.Id.ToString(), new NameValueCollection());
                        strBuilder.AppendFormat("{0} (RM#{1})", issue.Subject, issue.Id);
                        if (!string.IsNullOrEmpty(timeEntry.Comments))
                            strBuilder.Append(" - ");
                    }
                    
                    strBuilder.AppendFormat("{0} ({1} ч.);", timeEntry.Comments, timeEntry.Hours);
                    sw.WriteLine(strBuilder.ToString());

                }

            }

            var p = Process.Start("notepad.exe", path);

            if (p != null)
            {
                p.WaitForExit();
            }

            File.Delete(path);

            //Create a issue.
            //var newIssue = new Issue { Subject = "test", Project = new IdentifiableName { Id = 1 } };
            //manager.CreateObject(newIssue);
        }
    }
}
