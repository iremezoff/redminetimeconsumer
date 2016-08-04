using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using OutlookAddIn1.Properties;
using Office = Microsoft.Office.Core;

namespace OutlookAddIn1
{
    public partial class ReportCollectorForm : Form
    {
        private Dictionary<EmployeeItem, EmployeeReport> _reports;
        private Thread _requestThread;

        public ReportCollectorForm()
        {
            InitializeComponent();

            foreach (var employee in Settings.Default.Employees ?? new ArrayList())
            {
                var e = (EmployeeItem)employee;
                if (e.VacationEnd < DateTime.Now)
                {
                    e.InVacation = false;
                    e.VacationEnd = null;
                }
                bindingSource1.Add(employee);
            }
            subjectTailTextBox.DataBindings.Add("Text", Settings.Default, "SubjectTail");
            receiverTextBox.DataBindings.Add("Text", Settings.Default, "SendTo");
            redmineUriTextBox.DataBindings.Add("Text", Settings.Default, "RedmineUri");
            redmineTokenTextBox.DataBindings.Add("Text", Settings.Default, "RedmineToken");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CalendarColumn col = new CalendarColumn();
            col.Name = "VacationEnd";
            col.DataPropertyName = "VacationEnd";
            var unnecessaryColumn =
                employeeGridView.Columns.Cast<DataGridViewColumn>()
                    .FirstOrDefault(c => "VacationEnd".Equals(c.DataPropertyName));
            if(unnecessaryColumn!=null)
            this.employeeGridView.Columns.Remove(unnecessaryColumn);
            var column = this.employeeGridView.Columns.Add(col);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SelectNamesDialog Contact = (SelectNamesDialog)
            Globals.ThisAddIn.Application.Session.GetSelectNamesDialog();
            Contact.Display();
            Recipients tt = Contact.Recipients;
            foreach (Recipient item in tt)
            {
                receiverTextBox.Text += item.Name;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (_reports==null || !_reports.Any())
            {
                MessageBox.Show("Пусто", "Нет ничего", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            MailItem mailItem = (MailItem)
            Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
            mailItem.Subject = string.Format("Отчёт за {0:dd.MM.yyyy}. {1}", reportDateTimePicker.Value, Settings.Default.SubjectTail);
            mailItem.To = receiverTextBox.Text;
            //mailItem.Body = textBox3.Text+"\r\n";

            var strBuilder = new StringBuilder();

            foreach (var item in bindingSource1.List.OfType<EmployeeItem>())
            {
                strBuilder.AppendLine("--------------");
                strBuilder.AppendLine(item.Name);
                if (item.InVacation)
                    strBuilder.AppendFormat("В отпуске до {0:dd.MM.yyyy}\r\n", item.VacationEnd);
                else
                {
                    strBuilder.AppendLine(string.Join("\r\n", _reports[item].Items));
                    strBuilder.AppendLine();
                }
            }

            mailItem.Body = strBuilder.ToString();
            mailItem.Display(false);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (var item in bindingSource1.List.OfType<EmployeeItem>())
            {
                item.Hours = 0;
            }

            Settings.Default.Employees = new ArrayList(bindingSource1.List);

            Settings.Default.Save();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (_requestThread != null && _requestThread.IsAlive)
            {
                _requestThread.Abort();
                _requestThread.Interrupt();

                
                button3.Text = @"Запросить";
                return;
            }

            if (bindingSource1.List == null)
                return;

            SynchronizationContext ctx = new WindowsFormsSynchronizationContext();

            _requestThread = new Thread(RequestReport);

            _requestThread.Start(ctx);


            button3.Text = @"Долго! Я устал...";
        }

        private void RequestReport(object ctx)
        {
            var redmineService = new RedmineService(Settings.Default.RedmineUri,
                Settings.Default.RedmineToken);
            
            try
            {
            var goalEmployees = bindingSource1.List.Cast<EmployeeItem>().Where(i => !i.InVacation);

                _reports = redmineService.GetReports(goalEmployees, reportDateTimePicker.Value).ToDictionary(i => i.EmployeeItem, k => k);

                foreach (var item in goalEmployees)
                {
                    item.Hours = _reports[item].TotalHours;
                    item.Entries = _reports[item].Items.Count;
                }
            }
            catch (System.Exception ex)
            {
                
            }

            (ctx as SynchronizationContext).Post((obj) =>
            {
                bindingSource1.ResetBindings(false);

                button3.Text = @"Запросить";
            }, null);


        }
    }

    public class EmployeeItem
    {
        public string Name { get; set; }
        public bool InVacation { get; set; }
        public DateTime? VacationEnd { get; set; }
        public decimal Hours { get; set; }
        public int Entries { get; set; }
    }


    public class CalendarColumn : DataGridViewColumn
    {
        public CalendarColumn()
            : base(new CalendarCell())
        {
        }

        public override DataGridViewCell CellTemplate
        {
            get
            {
                return base.CellTemplate;
            }
            set
            {
                // Ensure that the cell used for the template is a CalendarCell. 
                if (value != null &&
                    !value.GetType().IsAssignableFrom(typeof(CalendarCell)))
                {
                    throw new InvalidCastException("Must be a CalendarCell");
                }
                base.CellTemplate = value;
            }
        }
    }

    public class CalendarCell : DataGridViewTextBoxCell
    {

        public CalendarCell()
            : base()
        {
            // Use the short date format. 
            this.Style.Format = "d";
        }

        public override void InitializeEditingControl(int rowIndex, object
            initialFormattedValue, DataGridViewCellStyle dataGridViewCellStyle)
        {
            // Set the value of the editing control to the current cell value. 
            base.InitializeEditingControl(rowIndex, initialFormattedValue,
                dataGridViewCellStyle);
            CalendarEditingControl ctl =
                DataGridView.EditingControl as CalendarEditingControl;
            // Use the default row value when Value property is null. 
            if (this.Value == null)
            {
                ctl.Value = (DateTime)this.DefaultNewRowValue;
            }
            else
            {
                ctl.Value = (DateTime)this.Value;
            }
        }

        public override Type EditType
        {
            get
            {
                // Return the type of the editing control that CalendarCell uses. 
                return typeof(CalendarEditingControl);
            }
        }

        public override Type ValueType
        {
            get
            {
                // Return the type of the value that CalendarCell contains. 

                return typeof(DateTime);
            }
        }

        public override object DefaultNewRowValue
        {
            get
            {
                // Use the current date and time as the default value. 
                return null;//return DateTime.Now;
            }
        }
    }

    class CalendarEditingControl : DateTimePicker, IDataGridViewEditingControl
    {
        DataGridView dataGridView;
        private bool valueChanged = false;
        int rowIndex;

        public CalendarEditingControl()
        {
            this.Format = DateTimePickerFormat.Short;
        }

        // Implements the IDataGridViewEditingControl.EditingControlFormattedValue  
        // property. 
        public object EditingControlFormattedValue
        {
            get
            {
                return this.Value.ToShortDateString();
            }
            set
            {
                if (value is String)
                {
                    try
                    {
                        // This will throw an exception of the string is  
                        // null, empty, or not in the format of a date. 
                        this.Value = DateTime.Parse((String)value);
                    }
                    catch
                    {
                        // In the case of an exception, just use the  
                        // default value so we're not left with a null 
                        // value. 
                        this.Value = DateTime.Now;
                    }
                }
            }
        }

        // Implements the  
        // IDataGridViewEditingControl.GetEditingControlFormattedValue method. 
        public object GetEditingControlFormattedValue(
            DataGridViewDataErrorContexts context)
        {
            return EditingControlFormattedValue;
        }

        // Implements the  
        // IDataGridViewEditingControl.ApplyCellStyleToEditingControl method. 
        public void ApplyCellStyleToEditingControl(
            DataGridViewCellStyle dataGridViewCellStyle)
        {
            this.Font = dataGridViewCellStyle.Font;
            this.CalendarForeColor = dataGridViewCellStyle.ForeColor;
            this.CalendarMonthBackground = dataGridViewCellStyle.BackColor;
        }

        // Implements the IDataGridViewEditingControl.EditingControlRowIndex  
        // property. 
        public int EditingControlRowIndex
        {
            get
            {
                return rowIndex;
            }
            set
            {
                rowIndex = value;
            }
        }

        // Implements the IDataGridViewEditingControl.EditingControlWantsInputKey  
        // method. 
        public bool EditingControlWantsInputKey(
            Keys key, bool dataGridViewWantsInputKey)
        {
            // Let the DateTimePicker handle the keys listed. 
            switch (key & Keys.KeyCode)
            {
                case Keys.Left:
                case Keys.Up:
                case Keys.Down:
                case Keys.Right:
                case Keys.Home:
                case Keys.End:
                case Keys.PageDown:
                case Keys.PageUp:
                    return true;
                default:
                    return !dataGridViewWantsInputKey;
            }
        }

        // Implements the IDataGridViewEditingControl.PrepareEditingControlForEdit  
        // method. 
        public void PrepareEditingControlForEdit(bool selectAll)
        {
            // No preparation needs to be done.
        }

        // Implements the IDataGridViewEditingControl 
        // .RepositionEditingControlOnValueChange property. 
        public bool RepositionEditingControlOnValueChange
        {
            get
            {
                return false;
            }
        }

        // Implements the IDataGridViewEditingControl 
        // .EditingControlDataGridView property. 
        public DataGridView EditingControlDataGridView
        {
            get
            {
                return dataGridView;
            }
            set
            {
                dataGridView = value;
            }
        }

        // Implements the IDataGridViewEditingControl 
        // .EditingControlValueChanged property. 
        public bool EditingControlValueChanged
        {
            get
            {
                return valueChanged;
            }
            set
            {
                valueChanged = value;
            }
        }

        // Implements the IDataGridViewEditingControl 
        // .EditingPanelCursor property. 
        public Cursor EditingPanelCursor
        {
            get
            {
                return base.Cursor;
            }
        }

        protected override void OnValueChanged(EventArgs eventargs)
        {
            // Notify the DataGridView that the contents of the cell 
            // have changed.
            valueChanged = true;
            this.EditingControlDataGridView.NotifyCurrentCellDirty(true);
            base.OnValueChanged(eventargs);
        }
    }


}
