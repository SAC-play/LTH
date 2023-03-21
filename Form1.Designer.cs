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
            TimeUnitUpDown.Location = new Point(133, 100);
            TimeUnitUpDown.Name = "TimeUnitUpDown";
            TimeUnitUpDown.ReadOnly = true;
            TimeUnitUpDown.Size = new Size(141, 23);
            TimeUnitUpDown.TabIndex = 2;
            TimeUnitUpDown.SelectedItemChanged += TimeUnitItemChanged;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
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
        public string[] m_str_time_unit = { "15 min", "30 min", "60 min" };
    }
}