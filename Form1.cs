using System.Windows.Forms;

namespace LTH
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void MouseClickOk(object sender, EventArgs e)
        {
            //work name input
        }

        private void textBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //work name input
            }
        }

        private void TimeUnitItemChanged(object sender, EventArgs e)
        {
            if (TimeUnitUpDown.SelectedItem.ToString() == m_str_time_unit[0])
            {
                //15 min
            }
            else if (TimeUnitUpDown.SelectedItem.ToString() == m_str_time_unit[1])
            {
                //30 min
            }
            else
            {
                //60 min
            }
        }
    }
}