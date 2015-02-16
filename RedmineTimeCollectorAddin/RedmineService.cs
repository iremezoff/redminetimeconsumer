using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Redmine.Net.Api;
using Redmine.Net.Api.Types;

namespace OutlookAddIn1
{
    class RedmineService
    {
        private readonly string _redmineUri;
        private readonly string _redmineToken;

        public RedmineService(string redmineUri, string redmineToken)
        {
            _redmineUri = redmineUri;
            _redmineToken = redmineToken;
        }

        public IEnumerable<EmployeeReport> GetReports(IEnumerable<EmployeeItem> employees, DateTime date)
        {
            var manager = new RedmineManager(_redmineUri, _redmineToken);

            var parameters = new NameValueCollection { { "spent_on", date.ToString("yyyy-MM-dd") }, { "limit", "1000" }, { "offset", "0" } };
            var timeEntries = manager.GetObjectList<TimeEntry>(parameters);

            foreach (var employeeItem in employees)
            {
                int i = 0;

                var report = new EmployeeReport() { EmployeeItem = employeeItem, Items = new List<string>() };
                
                var strBuilder = new StringBuilder();

                decimal totalHours = 0;

                foreach (var timeEntry in timeEntries.Where(e => e.User.Name.StartsWith(employeeItem.Name, StringComparison.CurrentCultureIgnoreCase)).OrderBy(e => e.CreatedOn)) //
                {
                    report.EmployeeItem.Name = timeEntry.User.Name;
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

                    report.Items.Add(strBuilder.ToString());

                    totalHours += timeEntry.Hours;
                }

                report.TotalHours = totalHours;

                yield return report;
            }
        }
    }

    class EmployeeReport
    {
        public EmployeeItem EmployeeItem { get; set; }

        public List<string> Items { get; set; }
        public decimal TotalHours { get; set; }
    }
}
