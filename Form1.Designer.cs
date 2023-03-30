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
            button2 = new Button();
            text_preiod = new TextBox();
            SuspendLayout();
            // 
            // textBox1
            // 
            textBox1.Location = new Point(133, 164);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(500, 23);
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
            TimeUnitUpDown.Items.Add("15 min");
            TimeUnitUpDown.Items.Add("30 min");
            TimeUnitUpDown.Items.Add("60 min");
            TimeUnitUpDown.Location = new Point(133, 106);
            TimeUnitUpDown.Name = "TimeUnitUpDown";
            TimeUnitUpDown.ReadOnly = true;
            TimeUnitUpDown.Size = new Size(141, 23);
            TimeUnitUpDown.TabIndex = 2;
            TimeUnitUpDown.SelectedItemChanged += TimeUnitItemChanged;
            // 
            // button2
            // 
            button2.Location = new Point(312, 223);
            button2.Name = "button2";
            button2.Size = new Size(181, 48);
            button2.TabIndex = 3;
            button2.Text = "Excel로 변환";
            button2.UseVisualStyleBackColor = true;
            // 
            // text_preiod
            // 
            text_preiod.BackColor = SystemColors.Control;
            text_preiod.Location = new Point(133, 135);
            text_preiod.Name = "text_preiod";
            text_preiod.ReadOnly = true;
            text_preiod.Size = new Size(256, 23);
            text_preiod.TabIndex = 4;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(text_preiod);
            Controls.Add(button2);
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
        public UInt16[] m_dTime_unit = { 15, 30, 60 };
        private Button button2;
        private TextBox text_preiod;
    }
}