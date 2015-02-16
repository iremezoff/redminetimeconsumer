namespace OutlookAddIn1
{
    partial class ReportCollectorForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.receiverTextBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.reportDateTimePicker = new System.Windows.Forms.DateTimePicker();
            this.button1 = new System.Windows.Forms.Button();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.button2 = new System.Windows.Forms.Button();
            this.employeeGridView = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewCheckBoxColumn1 = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.vacationTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Hours = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Entries = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.button3 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.redmineTokenTextBox = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.redmineUriTextBox = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.subjectTailTextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.contextMenuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.employeeGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // receiverTextBox
            // 
            this.receiverTextBox.Location = new System.Drawing.Point(119, 16);
            this.receiverTextBox.Name = "receiverTextBox";
            this.receiverTextBox.Size = new System.Drawing.Size(204, 20);
            this.receiverTextBox.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Получатель";
            // 
            // reportDateTimePicker
            // 
            this.reportDateTimePicker.CustomFormat = "dd.MM.yyyy";
            this.reportDateTimePicker.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.reportDateTimePicker.Location = new System.Drawing.Point(47, 12);
            this.reportDateTimePicker.Name = "reportDateTimePicker";
            this.reportDateTimePicker.Size = new System.Drawing.Size(158, 20);
            this.reportDateTimePicker.TabIndex = 3;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button1.Location = new System.Drawing.Point(89, 252);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(104, 52);
            this.button1.TabIndex = 6;
            this.button1.Text = "Сформировать";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1,
            this.toolStripMenuItem2});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(108, 48);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(107, 22);
            this.toolStripMenuItem1.Text = "Add";
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(107, 22);
            this.toolStripMenuItem2.Text = "Delete";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(329, 16);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(24, 20);
            this.button2.TabIndex = 21;
            this.button2.Text = "...";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // employeeGridView
            // 
            this.employeeGridView.AutoGenerateColumns = false;
            this.employeeGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.employeeGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewCheckBoxColumn1,
            this.vacationTextBoxColumn2,
            this.Hours,
            this.Entries});
            this.employeeGridView.DataSource = this.bindingSource1;
            this.employeeGridView.Location = new System.Drawing.Point(8, 38);
            this.employeeGridView.Name = "employeeGridView";
            this.employeeGridView.Size = new System.Drawing.Size(457, 208);
            this.employeeGridView.TabIndex = 22;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.DataPropertyName = "Name";
            this.dataGridViewTextBoxColumn1.HeaderText = "Name";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            // 
            // dataGridViewCheckBoxColumn1
            // 
            this.dataGridViewCheckBoxColumn1.DataPropertyName = "InVacation";
            this.dataGridViewCheckBoxColumn1.HeaderText = "InVacation";
            this.dataGridViewCheckBoxColumn1.Name = "dataGridViewCheckBoxColumn1";
            // 
            // vacationTextBoxColumn2
            // 
            this.vacationTextBoxColumn2.DataPropertyName = "VacationEnd";
            this.vacationTextBoxColumn2.HeaderText = "VacationEnd";
            this.vacationTextBoxColumn2.Name = "vacationTextBoxColumn2";
            // 
            // Hours
            // 
            this.Hours.DataPropertyName = "Hours";
            this.Hours.HeaderText = "Всего часов";
            this.Hours.Name = "Hours";
            this.Hours.ReadOnly = true;
            // 
            // Entries
            // 
            this.Entries.DataPropertyName = "Entries";
            this.Entries.HeaderText = "Пункты";
            this.Entries.Name = "Entries";
            this.Entries.ReadOnly = true;
            // 
            // bindingSource1
            // 
            this.bindingSource1.DataSource = typeof(OutlookAddIn1.EmployeeItem);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(8, 252);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 51);
            this.button3.TabIndex = 23;
            this.button3.Text = "Запросить";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.redmineTokenTextBox);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.redmineUriTextBox);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.subjectTailTextBox);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.receiverTextBox);
            this.groupBox1.Location = new System.Drawing.Point(471, 38);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(359, 208);
            this.groupBox1.TabIndex = 24;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Настройки";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 97);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(38, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "Токен";
            // 
            // redmineTokenTextBox
            // 
            this.redmineTokenTextBox.Location = new System.Drawing.Point(119, 94);
            this.redmineTokenTextBox.Name = "redmineTokenTextBox";
            this.redmineTokenTextBox.Size = new System.Drawing.Size(234, 20);
            this.redmineTokenTextBox.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 71);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(83, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Адрес Redmine";
            // 
            // redmineUriTextBox
            // 
            this.redmineUriTextBox.Location = new System.Drawing.Point(119, 68);
            this.redmineUriTextBox.Name = "redmineUriTextBox";
            this.redmineUriTextBox.Size = new System.Drawing.Size(234, 20);
            this.redmineUriTextBox.TabIndex = 5;
            this.redmineUriTextBox.Text = "http://redmine.ugsk.ru";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 45);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(107, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Дополнение к теме";
            // 
            // subjectTailTextBox
            // 
            this.subjectTailTextBox.Location = new System.Drawing.Point(119, 42);
            this.subjectTailTextBox.Name = "subjectTailTextBox";
            this.subjectTailTextBox.Size = new System.Drawing.Size(234, 20);
            this.subjectTailTextBox.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(33, 13);
            this.label2.TabIndex = 25;
            this.label2.Text = "Дата";
            // 
            // ReportCollectorForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(842, 324);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.employeeGridView);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.reportDateTimePicker);
            this.Name = "ReportCollectorForm";
            this.Text = "ReportCollectorForm";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.contextMenuStrip1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.employeeGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox receiverTextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker reportDateTimePicker;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.DataGridView employeeGridView;
        private System.Windows.Forms.BindingSource bindingSource1;
        private System.Windows.Forms.DataGridViewTextBoxColumn nameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn inVacationDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn vacationEndDataGridViewTextBoxColumn;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox subjectTailTextBox;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox redmineTokenTextBox;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox redmineUriTextBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewCheckBoxColumn dataGridViewCheckBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn vacationTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Hours;
        private System.Windows.Forms.DataGridViewTextBoxColumn Entries;
    }
}