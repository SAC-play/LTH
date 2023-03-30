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
            textBox1 = new TextBox();
            button1 = new Button();
            TimeUnitUpDown = new DomainUpDown();
            ExcelConvertButton = new Button();
            text_preiod = new TextBox();
            TimePeriodListView = new ListView();
            label1 = new Label();
            label_work_name = new Label();
            ListView_WorkName = new ListView();
            SuspendLayout();
            // 
            // textBox1
            // 
            textBox1.Location = new Point(327, 164);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(306, 23);
            textBox1.TabIndex = 0;
            textBox1.KeyDown += textBox_KeyDown;
            // 
            // button1
            // 
            button1.Location = new Point(648, 164);
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
            TimeUnitUpDown.Location = new Point(327, 106);
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
            ExcelConvertButton.Location = new Point(357, 213);
            ExcelConvertButton.Name = "ExcelConvertButton";
            ExcelConvertButton.Size = new Size(181, 48);
            ExcelConvertButton.TabIndex = 3;
            ExcelConvertButton.Text = "Excel로 변환";
            ExcelConvertButton.UseVisualStyleBackColor = true;
            ExcelConvertButton.MouseClick += ExcelConvertButtonClick;
            // 
            // text_preiod
            // 
            text_preiod.BackColor = SystemColors.Control;
            text_preiod.Location = new Point(327, 135);
            text_preiod.Name = "text_preiod";
            text_preiod.ReadOnly = true;
            text_preiod.Size = new Size(141, 23);
            text_preiod.TabIndex = 4;
            // 
            // TimePeriodListView
            // 
            TimePeriodListView.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            TimePeriodListView.Location = new Point(12, 106);
            TimePeriodListView.MultiSelect = false;
            TimePeriodListView.Name = "TimePeriodListView";
            TimePeriodListView.Size = new Size(111, 198);
            TimePeriodListView.TabIndex = 5;
            TimePeriodListView.UseCompatibleStateImageBehavior = false;
            TimePeriodListView.View = View.List;
            TimePeriodListView.SelectedIndexChanged += TPListViewSelectedIndexChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Enabled = false;
            label1.Font = new Font("맑은 고딕", 10F, FontStyle.Bold, GraphicsUnit.Point);
            label1.Location = new Point(43, 84);
            label1.Name = "label1";
            label1.Size = new Size(42, 19);
            label1.TabIndex = 6;
            label1.Text = "시 간";
            // 
            // label_work_name
            // 
            label_work_name.AutoSize = true;
            label_work_name.Enabled = false;
            label_work_name.Font = new Font("맑은 고딕", 10F, FontStyle.Bold, GraphicsUnit.Point);
            label_work_name.Location = new Point(168, 84);
            label_work_name.Name = "label_work_name";
            label_work_name.Size = new Size(80, 19);
            label_work_name.TabIndex = 8;
            label_work_name.Text = "업 무 내 용";
            // 
            // ListView_WorkName
            // 
            ListView_WorkName.Location = new Point(126, 106);
            ListView_WorkName.Name = "ListView_WorkName";
            ListView_WorkName.Size = new Size(180, 198);
            ListView_WorkName.TabIndex = 7;
            ListView_WorkName.UseCompatibleStateImageBehavior = false;
            ListView_WorkName.View = View.List;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(label_work_name);
            Controls.Add(ListView_WorkName);
            Controls.Add(label1);
            Controls.Add(TimePeriodListView);
            Controls.Add(text_preiod);
            Controls.Add(ExcelConvertButton);
            Controls.Add(TimeUnitUpDown);
            Controls.Add(button1);
            Controls.Add(textBox1);
            Name = "Form1";
            Text = "LetsGoHaeNi";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox textBox1;
        private Button button1;
        private DomainUpDown TimeUnitUpDown;
        public UInt16[] m_dTime_unit = { 10, 15, 30, 60 };
        private Button ExcelConvertButton;
        private TextBox text_preiod;
        private ListView TimePeriodListView;
        private Label label1;
        private Label label_work_name;
        private ListView ListView_WorkName;
    }
}