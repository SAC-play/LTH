using System.Timers;
using System.Windows.Forms;

namespace LTH
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            var mon_dt = m_sht_obj.find_monday_date();

            m_excel_io.create_file("업무일지_" + mon_dt.Month.ToString("00") + mon_dt.Day.ToString("00") + ".xlsx");

            m_sht_obj.set_time_unit((double)15);
            m_sht_obj.calculate_start_time();
            m_sht_obj.add_elapsed_handler(on_elapsed);

            //m_sht_obj.EndTime = DateTime.Now.AddMinutes(1);

            var begin_dt = m_sht_obj.BeginTime;
            var end_dt = m_sht_obj.EndTime;

            text_preiod.Text = begin_dt.Hour.ToString("00") + ":" + begin_dt.Minute.ToString("00")+" ~ "+ end_dt.Hour.ToString("00") + ":" + end_dt.Minute.ToString("00");

            m_sht_obj.start_timer();
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
            if (TimeUnitUpDown.SelectedItem.ToString() == m_dTime_unit[0].ToString() + " min")
            {
                //15 min
                m_sht_obj.set_time_unit((double)(m_dTime_unit[0]));
            }
            else if (TimeUnitUpDown.SelectedItem.ToString() == m_dTime_unit[1].ToString() + " min")
            {
                //30 min
                m_sht_obj.set_time_unit((double)(m_dTime_unit[1]));
            }
            else
            {
                //60 min
                m_sht_obj.set_time_unit((double)60);
            }
        }

        delegate void TimerEventFiredDelegate(string text);

        public void on_elapsed(Object source, EventArgs e)
        {
            var sht_obj = source as SaveHerTime;

            //MessageBox.Show("[on_elapsed]\nbegin time : "+sht_obj.BeginTime.ToString() + "\nend time : "+sht_obj.EndTime.ToString());

            var begin_dt = sht_obj.EndTime;
            var end_dt = sht_obj.EndTime.AddMinutes(sht_obj.TimeUnit);

            string preiod_text = begin_dt.Hour.ToString("00") + ":" + begin_dt.Minute.ToString("00") + " ~ " + end_dt.Hour.ToString("00") + ":" + end_dt.Minute.ToString("00"); ;

            BeginInvoke(new TimerEventFiredDelegate(ui_work), preiod_text);

            sht_obj.ChangeBeginTIme = true;
        }

        public void ui_work(string period_text)
        {
            //below statement is caused error.
            text_preiod.Text = period_text;
        }

        private Excel_io m_excel_io = new Excel_io();
        private SaveHerTime m_sht_obj = new SaveHerTime();
    }
}