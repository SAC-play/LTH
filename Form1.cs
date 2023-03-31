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

            m_excel_io.create_file("诀公老瘤_" + mon_dt.Month.ToString("00") + mon_dt.Day.ToString("00") + ".xlsx");

#if true //For test, if you want to end time set 1 min later, then false.
            m_sht_obj.set_time_unit((double)10);
            m_sht_obj.calculate_start_time();
#else
            m_sht_obj.set_time_unit((double)1);
            m_sht_obj.EndTime = DateTime.Now.AddMinutes(1);
#endif
            m_sht_obj.add_elapsed_handler(on_elapsed);

            var begin_dt = m_sht_obj.BeginTime;
            var end_dt = m_sht_obj.EndTime;

            string str_period = begin_dt.Hour.ToString("00") + ":" + begin_dt.Minute.ToString("00") + " ~ " + end_dt.Hour.ToString("00") + ":" + end_dt.Minute.ToString("00");

            text_preiod.Text = str_period;

            m_sht_obj.start_timer();

            var dt_now = DateTime.Now;

            if (dt_now.DayOfWeek == DayOfWeek.Monday)
            {
                m_cExcel_stdTime_column = 'B';
            }
            else if (dt_now.DayOfWeek == DayOfWeek.Tuesday)
            {
                m_cExcel_stdTime_column = 'D';
            }
            else if (dt_now.DayOfWeek == DayOfWeek.Wednesday)
            {
                m_cExcel_stdTime_column = 'F';
            }
            else if (dt_now.DayOfWeek == DayOfWeek.Thursday)
            {
                m_cExcel_stdTime_column = 'H';
            }
            else if (dt_now.DayOfWeek == DayOfWeek.Friday)
            {
                m_cExcel_stdTime_column = 'J';
            }
            else if (dt_now.DayOfWeek == DayOfWeek.Saturday)
            {
                m_cExcel_stdTime_column = 'L';
            }
            else //Sunday
            {
                m_cExcel_stdTime_column = 'N';
            }

            m_cExcel_stdCtxt_column = (char)((int)m_cExcel_stdTime_column + 1);
            m_nExcel_std_row = 3;

            m_excel_io.set_data(1, "A", "诀公老瘤", Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen);
            m_excel_io.set_data((m_nExcel_std_row - 1), m_cExcel_stdCtxt_column.ToString(), DateTime.Now.ToString("MM岿 dd老 ddd"), Microsoft.Office.Interop.Excel.XlRgbColor.rgbDarkGray);

            m_dt_beginTime = m_sht_obj.BeginTime;
            m_dt_endTime = m_sht_obj.EndTime;
            m_dt_future_endTime = m_dt_endTime.AddMinutes(m_sht_obj.TimeUnit);
        }

        private void MouseClickOk(object sender, EventArgs e)
        {
            //work name input
            if (textBox1.Text.Length != 0)
            {
                m_bInputStringLeastOnceInPeriod = true;

                //add context into save her time dictionary
                m_excel_io.set_data(m_nExcel_std_row, m_cExcel_stdCtxt_column.ToString(), textBox1.Text.ToString());

                string time_dict_key = m_nExcel_std_row.ToString() + m_cExcel_stdTime_column.ToString();

                //MessageBox.Show(textBox1.Text);
                //if dict key is not existed in DictData, then save data and update time period list view.
                if (!m_excel_io.DictData.ContainsKey(time_dict_key))
                {
                    //make period text
                    var begin_dt = m_dt_beginTime;
                    var end_dt = m_dt_endTime;

                    string period_text = begin_dt.Hour.ToString("00") + ":" + begin_dt.Minute.ToString("00") + " ~ " + end_dt.Hour.ToString("00") + ":" + end_dt.Minute.ToString("00");

                    TimePeriodListView.Items.Add(period_text);

                    m_sht_obj.ChangeBeginTIme = true;

                    m_excel_io.set_data(m_nExcel_std_row, m_cExcel_stdTime_column.ToString(), period_text, Microsoft.Office.Interop.Excel.XlRgbColor.rgbLightGray);
                }

                //list view update
                //make context list key
                string context_dict_key = m_nExcel_std_row.ToString() + m_cExcel_stdCtxt_column.ToString();

                if (!m_list_context_keys.Contains(context_dict_key))
                {
                    m_list_context_keys.Add(context_dict_key);
                }

                if (TimePeriodListView.Items.Count == 1)
                {
                    selected_idx = 0;
                    listview_workname_update(m_list_context_keys[0]);
                }
                else
                {
                    if (selected_idx != (TimePeriodListView.Items.Count - 1))
                    {
                        TimePeriodListView.Items[TimePeriodListView.Items.Count - 1].Selected = true;
                    }
                    else
                    {
                        listview_workname_update(m_list_context_keys[selected_idx]);
                    }
                }

                /*
                TimePeriodListView.BeginUpdate();
                TimePeriodListView.Items.Add(text_preiod);
                TimePeriodListView.Items[TimePeriodListView.Items.Count - 1].Selected = true;
                TimePeriodListView.EndUpdate();

                ListView_WorkName.Items.Add(textBox1.Text.ToString());
                */

                textBox1.Text = "";
            }
        }

        private void textBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;

                //work name input
                if (textBox1.Text.Length != 0)
                {
                    m_bInputStringLeastOnceInPeriod = true;

                    //add context into save her time dictionary
                    m_excel_io.set_data(m_nExcel_std_row, m_cExcel_stdCtxt_column.ToString(), textBox1.Text.ToString());

                    string time_dict_key = m_nExcel_std_row.ToString() + m_cExcel_stdTime_column.ToString();

                    //MessageBox.Show(textBox1.Text);
                    //if dict key is not existed in DictData, then save data and update time period list view.
                    if (!m_excel_io.DictData.ContainsKey(time_dict_key))
                    {
                        //make period text
                        var begin_dt = m_dt_beginTime;
                        var end_dt = m_dt_endTime;

                        string period_text = begin_dt.Hour.ToString("00") + ":" + begin_dt.Minute.ToString("00") + " ~ " + end_dt.Hour.ToString("00") + ":" + end_dt.Minute.ToString("00");

                        TimePeriodListView.Items.Add(period_text);

                        m_sht_obj.ChangeBeginTIme = true;

                        m_excel_io.set_data(m_nExcel_std_row, m_cExcel_stdTime_column.ToString(), period_text, Microsoft.Office.Interop.Excel.XlRgbColor.rgbLightGray);
                    }

                    //list view update
                    //make context list key
                    string context_dict_key = m_nExcel_std_row.ToString() + m_cExcel_stdCtxt_column.ToString();

                    if(!m_list_context_keys.Contains(context_dict_key))
                    {
                        m_list_context_keys.Add(context_dict_key);
                    }

                    if(TimePeriodListView.Items.Count == 1)
                    {
                        selected_idx = 0;
                        listview_workname_update(m_list_context_keys[0]);
                    }
                    else
                    {
                        if (selected_idx != (TimePeriodListView.Items.Count-1))
                        {
                            TimePeriodListView.Items[TimePeriodListView.Items.Count - 1].Selected = true;
                        }
                        else
                        {
                            listview_workname_update(m_list_context_keys[selected_idx]);
                        }
                    }

                    /*
                    TimePeriodListView.BeginUpdate();
                    TimePeriodListView.Items.Add(text_preiod);
                    TimePeriodListView.Items[TimePeriodListView.Items.Count - 1].Selected = true;
                    TimePeriodListView.EndUpdate();

                    ListView_WorkName.Items.Add(textBox1.Text.ToString());
                    */

                    textBox1.Text = "";
                }
            }
        }

        private void listview_workname_update(string dict_key)
        {
            ListView_WorkName.BeginUpdate();
            ListView_WorkName.Clear();

            foreach (var item in m_excel_io.DictData[dict_key].list_datas)
            {
                ListView_WorkName.Items.Add(item);
            }

            ListView_WorkName.EndUpdate();
        }

        private void TimeUnitItemChanged(object sender, EventArgs e)
        {
            bool bEndTimeChanged = false;
            if (TimeUnitUpDown.SelectedItem.ToString() == m_dTime_unit[0].ToString() + " min")
            {
                //10 min
                bEndTimeChanged = m_sht_obj.set_time_unit((double)(m_dTime_unit[0]));
            }
            else if (TimeUnitUpDown.SelectedItem.ToString() == m_dTime_unit[1].ToString() + " min")
            {
                //15 min
                bEndTimeChanged = m_sht_obj.set_time_unit((double)(m_dTime_unit[1]));
            }
            else if (TimeUnitUpDown.SelectedItem.ToString() == m_dTime_unit[2].ToString() + " min")
            {
                //30 min
                bEndTimeChanged = m_sht_obj.set_time_unit((double)(m_dTime_unit[2]));
            }
            else if (TimeUnitUpDown.SelectedItem.ToString() == m_dTime_unit[3].ToString() + " min")
            {
                //60 min
                bEndTimeChanged = m_sht_obj.set_time_unit((double)(m_dTime_unit[3]));
            }
            else
            {
                // exception
            }

            if (bEndTimeChanged)
            {
                var begin_dt = m_sht_obj.BeginTime;
                var end_dt = m_sht_obj.EndTime;

                text_preiod.Text = begin_dt.Hour.ToString("00") + ":" + begin_dt.Minute.ToString("00") + " ~ " + end_dt.Hour.ToString("00") + ":" + end_dt.Minute.ToString("00");
            }
        }

        delegate void TimerEventFiredDelegate(string period_text);

        public void on_elapsed(Object source, EventArgs e)
        {
            //var sht_obj = source as SaveHerTime
            var param_args = e as SaveHerTime.SaveHerTimeEventArgs;
            string period_text = "";

            m_dt_beginTime = param_args.BeginTime;
            m_dt_endTime = param_args.EndTime;
            m_dt_future_endTime = param_args.FutureEndTime;

            period_text = m_dt_beginTime.Hour.ToString("00") + ":" + m_dt_beginTime.Minute.ToString("00") + " ~ " + m_dt_endTime.Hour.ToString("00") + ":" + m_dt_endTime.Minute.ToString("00");

            if (m_bInputStringLeastOnceInPeriod)
            {
                m_nExcel_std_row++;

                m_sht_obj.ChangeBeginTIme = true;
            }

            m_bInputStringLeastOnceInPeriod = false;

            BeginInvoke(new TimerEventFiredDelegate(ui_work), period_text);
        }

        public void ui_work(string period_text)
        {
            text_preiod.Text = period_text;
        }

        private void ExcelConvertButtonClick(object sender, MouseEventArgs e)
        {
            m_excel_io.sync_data();
        }

        private void TPListViewSelectedIndexChanged(object sender, EventArgs e)
        {
            if((TimePeriodListView.SelectedItems.Count > 0))
            {
                listview_workname_update(m_list_context_keys[TimePeriodListView.SelectedIndices[0]]);

                //MessageBox.Show("TPListViewSelectedIndexChanged index : " + selected_idx.ToString());

                selected_idx = TimePeriodListView.SelectedIndices[0];
            }
        }

        private Excel_io m_excel_io = new Excel_io();
        private SaveHerTime m_sht_obj = new SaveHerTime();
        private char m_cExcel_stdTime_column = 'B';
        private char m_cExcel_stdCtxt_column = 'C';
        private int m_nExcel_std_row = 0;
        private bool m_bInputStringLeastOnceInPeriod = true;
        private List<string> m_list_context_keys = new List<string>();
        private int selected_idx = -1;
        private DateTime m_dt_beginTime;
        private DateTime m_dt_endTime;
        private DateTime m_dt_future_endTime;
    }
}