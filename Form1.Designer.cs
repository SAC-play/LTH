namespace LTH
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            textBox_workname = new TextBox();
            button1 = new Button();
            TimeUnitUpDown = new DomainUpDown();
            ExcelConvertButton = new Button();
            textBox_period = new TextBox();
            HowToWrite = new Label();
            label1 = new Label();
            TimePeriodListView = new ListView();
            label_work_name = new Label();
            ListView_WorkName = new ListView();
            PreviewGroup = new GroupBox();
            PreviewGroup.SuspendLayout();
            SuspendLayout();
            // 
            // textBox_workname
            // 
            textBox_workname.Location = new Point(341, 189);
            textBox_workname.Name = "textBox_workname";
            textBox_workname.Size = new Size(306, 23);
            textBox_workname.TabIndex = 0;
            textBox_workname.KeyDown += textBox_KeyDown;
            // 
            // button1
            // 
            button1.Location = new Point(664, 189);
            button1.Name = "button1";
            button1.Size = new Size(101, 23);
            button1.TabIndex = 1;
            button1.Text = "입  력";
            button1.UseVisualStyleBackColor = true;
            button1.Click += MouseClickOk;
            // 
            // TimeUnitUpDown
            // 
            TimeUnitUpDown.BackColor = SystemColors.Control;
            TimeUnitUpDown.Items.Add("60 min");
            TimeUnitUpDown.Items.Add("30 min");
            TimeUnitUpDown.Items.Add("15 min");
            TimeUnitUpDown.Items.Add("10 min");
            TimeUnitUpDown.Location = new Point(341, 106);
            TimeUnitUpDown.Name = "TimeUnitUpDown";
            TimeUnitUpDown.ReadOnly = true;
            TimeUnitUpDown.Size = new Size(141, 23);
            TimeUnitUpDown.TabIndex = 2;
            TimeUnitUpDown.Text = "10 min";
            TimeUnitUpDown.Wrap = true;
            TimeUnitUpDown.SelectedItemChanged += TimeUnitItemChanged;
            // 
            // ExcelConvertButton
            // 
            ExcelConvertButton.Location = new Point(400, 256);
            ExcelConvertButton.Name = "ExcelConvertButton";
            ExcelConvertButton.Size = new Size(181, 48);
            ExcelConvertButton.TabIndex = 3;
            ExcelConvertButton.Text = "Excel로 변환";
            ExcelConvertButton.UseVisualStyleBackColor = true;
            ExcelConvertButton.MouseClick += ExcelConvertButtonClick;
            // 
            // textBox_period
            // 
            textBox_period.BackColor = SystemColors.Control;
            textBox_period.Location = new Point(341, 135);
            textBox_period.Name = "textBox_period";
            textBox_period.ReadOnly = true;
            textBox_period.Size = new Size(141, 23);
            textBox_period.TabIndex = 4;
            // 
            // HowToWrite
            // 
            HowToWrite.AutoSize = true;
            HowToWrite.Location = new Point(341, 171);
            HowToWrite.Name = "HowToWrite";
            HowToWrite.Size = new Size(254, 15);
            HowToWrite.TabIndex = 9;
            HowToWrite.Text = "작성방법( 회사명_항목명_작업명_세부작업명 )";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Enabled = false;
            label1.Font = new Font("맑은 고딕", 10F, FontStyle.Bold, GraphicsUnit.Point);
            label1.Location = new Point(48, 23);
            label1.Name = "label1";
            label1.Size = new Size(42, 19);
            label1.TabIndex = 6;
            label1.Text = "시 간";
            // 
            // TimePeriodListView
            // 
            TimePeriodListView.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            TimePeriodListView.Location = new Point(11, 45);
            TimePeriodListView.MultiSelect = false;
            TimePeriodListView.Name = "TimePeriodListView";
            TimePeriodListView.Size = new Size(111, 198);
            TimePeriodListView.TabIndex = 5;
            TimePeriodListView.UseCompatibleStateImageBehavior = false;
            TimePeriodListView.View = View.List;
            TimePeriodListView.SelectedIndexChanged += TPListViewSelectedIndexChanged;
            // 
            // label_work_name
            // 
            label_work_name.AutoSize = true;
            label_work_name.Enabled = false;
            label_work_name.Font = new Font("맑은 고딕", 10F, FontStyle.Bold, GraphicsUnit.Point);
            label_work_name.Location = new Point(184, 23);
            label_work_name.Name = "label_work_name";
            label_work_name.Size = new Size(80, 19);
            label_work_name.TabIndex = 8;
            label_work_name.Text = "업 무 내 용";
            // 
            // ListView_WorkName
            // 
            ListView_WorkName.Location = new Point(133, 45);
            ListView_WorkName.Name = "ListView_WorkName";
            ListView_WorkName.Size = new Size(180, 198);
            ListView_WorkName.TabIndex = 7;
            ListView_WorkName.UseCompatibleStateImageBehavior = false;
            ListView_WorkName.View = View.List;
            // 
            // PreviewGroup
            // 
            PreviewGroup.Controls.Add(ListView_WorkName);
            PreviewGroup.Controls.Add(label_work_name);
            PreviewGroup.Controls.Add(TimePeriodListView);
            PreviewGroup.Controls.Add(label1);
            PreviewGroup.Font = new Font("맑은 고딕", 10F, FontStyle.Bold, GraphicsUnit.Point);
            PreviewGroup.Location = new Point(12, 79);
            PreviewGroup.Name = "PreviewGroup";
            PreviewGroup.Size = new Size(323, 263);
            PreviewGroup.TabIndex = 10;
            PreviewGroup.TabStop = false;
            PreviewGroup.Text = "Preview";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(HowToWrite);
            Controls.Add(textBox_period);
            Controls.Add(ExcelConvertButton);
            Controls.Add(TimeUnitUpDown);
            Controls.Add(button1);
            Controls.Add(textBox_workname);
            Controls.Add(PreviewGroup);
            Name = "Form1";
            Text = "LetsGoHaeNi";
            PreviewGroup.ResumeLayout(false);
            PreviewGroup.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox textBox_workname;
        private Button button1;
        private DomainUpDown TimeUnitUpDown;
        public UInt16[] m_dTime_unit = { 10, 15, 30, 60 };
        private Button ExcelConvertButton;
        private TextBox textBox_period;
        private Label HowToWrite;
        private Label label1;
        private ListView TimePeriodListView;
        private Label label_work_name;
        private ListView ListView_WorkName;
        private GroupBox PreviewGroup;
    }
}