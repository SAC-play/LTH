using System.Text.Json.Nodes;
using System.Timers;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace LTH
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            var mon_dt = m_sht_obj.find_monday_date();

            m_excel_io.mySheet = mon_dt.Month.ToString("00") + mon_dt.Day.ToString("00");
            m_excel_io.create_file("업무일지_" + mon_dt.Month.ToString("00") + "월.xlsx");

            initialize_cell_position();

            initialize_timer();

            //timer period text box initialize
            var begin_dt = m_sht_obj.BeginTime;
            var end_dt = m_sht_obj.EndTime;

            string str_period = begin_dt.Hour.ToString("00") + ":" + begin_dt.Minute.ToString("00") + " ~ " + end_dt.Hour.ToString("00") + ":" + end_dt.Minute.ToString("00");

            textBox_period.Text = str_period;

            var list_keys = m_excel_io.DictData.Keys.ToList();
            var excel_dict_data = m_excel_io.DictData;
            list_keys.Sort();

            foreach (var key in list_keys)
            {
                // time period key
                if ((Int32.Parse(key[0].ToString()) >= 3) && (key[1] == m_cExcel_stdTime_column))
                {
                    TimePeriodListView.Items.Add(excel_dict_data[key].list_datas[0]);
                }

                //context key
                if ((Int32.Parse(key[0].ToString()) >= 3) && (key[1] == m_cExcel_stdCtxt_column))
                {
                    m_list_context_keys.Add(key);
                }
            }

            if (TimePeriodListView.Items.Count > 0)
            {
                TimePeriodListView.Items[TimePeriodListView.Items.Count - 1].Selected = true;
            }
        }

        ~Form1()
        {
            m_timer.Stop();
        }

        private void handle_work_name()
        {
            m_bInputStringLeastOnceInPeriod = true;

            var begin_dt = m_dt_beginTime;
            var end_dt = m_dt_endTime;

            //make period text

            //lunch time calculate
            if ((begin_dt.Hour < 12) && (end_dt.Hour >= 13))
            {
                //before lunch time
                string time_dict_key = m_nExcel_std_row.ToString() + m_cExcel_stdTime_column.ToString();

                if (!m_excel_io.DictData.ContainsKey(time_dict_key))
                {
                    string period_text = begin_dt.Hour.ToString("00") + ":" + begin_dt.Minute.ToString("00") + " ~ 12:00";

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

                //add context into save her time dictionary
                m_excel_io.set_data(m_nExcel_std_row, m_cExcel_stdCtxt_column.ToString(), textBox_workname.Text.ToString());
                m_nExcel_std_row++;

                //after lunch time
                time_dict_key = m_nExcel_std_row.ToString() + m_cExcel_stdTime_column.ToString();

                if (!m_excel_io.DictData.ContainsKey(time_dict_key))
                {
                    string period_text = "13:00 ~ " + end_dt.Hour.ToString("00") + ":" + end_dt.Minute.ToString("00");

                    TimePeriodListView.Items.Add(period_text);

                    m_sht_obj.ChangeBeginTIme = true;

                    m_excel_io.set_data(m_nExcel_std_row, m_cExcel_stdTime_column.ToString(), period_text, Microsoft.Office.Interop.Excel.XlRgbColor.rgbLightGray);
                }
                //list view update
                //make context list key
                context_dict_key = m_nExcel_std_row.ToString() + m_cExcel_stdCtxt_column.ToString();

                if (!m_list_context_keys.Contains(context_dict_key))
                {
                    m_list_context_keys.Add(context_dict_key);
                }

                //add context into save her time dictionary
                m_excel_io.set_data(m_nExcel_std_row, m_cExcel_stdCtxt_column.ToString(), textBox_workname.Text.ToString());


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

                m_sht_obj.BeginTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 13, 0, 0);
                m_dt_beginTime = m_sht_obj.BeginTime;
            }
            //dinner time
            else if ((begin_dt.Hour < 18) && (end_dt.Hour >= 19))
            {
                //before dinner time
                string time_dict_key = m_nExcel_std_row.ToString() + m_cExcel_stdTime_column.ToString();

                if (!m_excel_io.DictData.ContainsKey(time_dict_key))
                {
                    string period_text = begin_dt.Hour.ToString("00") + ":" + begin_dt.Minute.ToString("00") + " ~ 18:00";

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

                //add context into save her time dictionary
                m_excel_io.set_data(m_nExcel_std_row, m_cExcel_stdCtxt_column.ToString(), textBox_workname.Text.ToString());
                m_nExcel_std_row++;

                //after lunch time
                time_dict_key = m_nExcel_std_row.ToString() + m_cExcel_stdTime_column.ToString();

                if (!m_excel_io.DictData.ContainsKey(time_dict_key))
                {
                    string period_text = "19:00 ~ " + end_dt.Hour.ToString("00") + ":" + end_dt.Minute.ToString("00");

                    TimePeriodListView.Items.Add(period_text);

                    m_sht_obj.ChangeBeginTIme = true;

                    m_excel_io.set_data(m_nExcel_std_row, m_cExcel_stdTime_column.ToString(), period_text, Microsoft.Office.Interop.Excel.XlRgbColor.rgbLightGray);
                }
                //list view update
                //make context list key
                context_dict_key = m_nExcel_std_row.ToString() + m_cExcel_stdCtxt_column.ToString();

                if (!m_list_context_keys.Contains(context_dict_key))
                {
                    m_list_context_keys.Add(context_dict_key);
                }

                //add context into save her time dictionary
                m_excel_io.set_data(m_nExcel_std_row, m_cExcel_stdCtxt_column.ToString(), textBox_workname.Text.ToString());


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

                m_sht_obj.BeginTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 19, 0, 0);
                m_dt_beginTime = m_sht_obj.BeginTime;
            }
            else
            {
                //add context into save her time dictionary
                m_excel_io.set_data(m_nExcel_std_row, m_cExcel_stdCtxt_column.ToString(), textBox_workname.Text.ToString());

                string time_dict_key = m_nExcel_std_row.ToString() + m_cExcel_stdTime_column.ToString();

                //MessageBox.Show(textBox1.Text);
                //if dict key is not existed in DictData, then save data and update time period list view.
                if (!m_excel_io.DictData.ContainsKey(time_dict_key))
                {
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
            }

            textBox_workname.Text = "";
        }

        private void MouseClickOk(object sender, EventArgs e)
        {
            //work name input
            if (textBox_workname.Text.Length != 0)
            {
                handle_work_name();
            }
        }

        private void textBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;

                handle_work_name();
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

                textBox_period.Text = begin_dt.Hour.ToString("00") + ":" + begin_dt.Minute.ToString("00") + " ~ " + end_dt.Hour.ToString("00") + ":" + end_dt.Minute.ToString("00");
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
            textBox_period.Text = period_text;
        }

        private void ExcelConvertButtonClick(object sender, MouseEventArgs e)
        {
            m_excel_io.sync_data();
            m_excel_io.clear_dictionary();
            m_list_context_keys.Clear();

            TimePeriodListView.Clear();
            ListView_WorkName.Clear();

            string auto_save_file_path_name = AppDomain.CurrentDomain.BaseDirectory.ToString() + m_auto_save_file_name;

            if (System.IO.File.Exists(auto_save_file_path_name))
            {
                File.Delete(auto_save_file_path_name);
            }
        }

        private void TPListViewSelectedIndexChanged(object sender, EventArgs e)
        {
            if ((TimePeriodListView.SelectedItems.Count > 0))
            {
                listview_workname_update(m_list_context_keys[TimePeriodListView.SelectedIndices[0]]);

                //MessageBox.Show("TPListViewSelectedIndexChanged index : " + selected_idx.ToString());

                selected_idx = TimePeriodListView.SelectedIndices[0];
            }
        }

        private void auto_save(Object source, ElapsedEventArgs e)
        {
            var dict_data = m_excel_io.DictData;

            if (dict_data.Count == 0)
            {
                return;
            }

            JObject json_temp_data = new JObject();

            foreach (var item in m_excel_io.DictData)
            {
                string[] str_array = item.Value.list_datas.ToArray();

                JObject item_obj = new JObject(
                    new JProperty("data", str_array),
                    new JProperty("color", (int)item.Value.rgb_color)
                );

                json_temp_data.Add(item.Key, item_obj);
            }

            {
                string auto_save_file_path_name = AppDomain.CurrentDomain.BaseDirectory.ToString() + m_auto_save_file_name;

                File.WriteAllText(auto_save_file_path_name, json_temp_data.ToString());
            }
        }

        private void initialize_timer()
        {

#if true //For test, if you want to end time set 1 min later, then false.
            if (m_nExcel_std_row > 3)
            {
                string time_dict_key = (m_nExcel_std_row - 1).ToString() + m_cExcel_stdTime_column.ToString();
                string str_time = m_excel_io.DictData[time_dict_key].list_datas[0].ToString();
                string begin_hour = str_time.Substring(0, 2);
                string begin_min = str_time.Substring(3, 2);
                string end_hour = str_time.Substring(8, 2);
                string end_min = str_time.Substring(11, 2);

                DateTime my_begin_dt = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, Int32.Parse(begin_hour), Int32.Parse(begin_min), 0);
                DateTime my_end_dt = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, Int32.Parse(end_hour), Int32.Parse(end_min), 0);

                if (!m_sht_obj.calculate_start_time(my_begin_dt, my_end_dt))
                {
                    m_nExcel_std_row--;
                }
            }
            else
            {
                m_sht_obj.set_time_unit((double)10);
                m_sht_obj.calculate_start_time();
            }
#else
            m_sht_obj.set_time_unit((double)1);
            //m_sht_obj.EndTime = DateTime.Now.AddMinutes(1);
            m_sht_obj.BeginTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 9, 10, 0);
            m_sht_obj.EndTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 13, 30, 0);
#endif
            m_sht_obj.add_elapsed_handler(on_elapsed);

            m_sht_obj.start_timer();

            m_dt_beginTime = m_sht_obj.BeginTime;
            m_dt_endTime = m_sht_obj.EndTime;

            m_timer.Elapsed += auto_save;
            m_timer.Start();
        }

        private void initialize_cell_position()
        {
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
            string auto_save_file_path_name = AppDomain.CurrentDomain.BaseDirectory.ToString() + m_auto_save_file_name;

            if (System.IO.File.Exists(auto_save_file_path_name))
            {
                StreamReader file = File.OpenText(auto_save_file_path_name);
                JsonTextReader textReader = new JsonTextReader(file);
                JObject json_object = (JObject)JToken.ReadFrom(textReader);

                file.Close();

                int pivot_row = 0;

                foreach (JProperty obj in json_object.Properties())
                {
                    string dict_key = obj.Name;

                    int nRow = Int32.Parse(dict_key[0].ToString());

                    Microsoft.Office.Interop.Excel.XlRgbColor color = obj.Value["color"].ToObject<Microsoft.Office.Interop.Excel.XlRgbColor>();

                    foreach (var item in (JArray)obj.Value["data"])
                    {
                        m_excel_io.set_data(nRow, dict_key[1].ToString(), item.ToString(), color);
                    }

                    if (pivot_row < nRow)
                    {
                        pivot_row = nRow;
                        m_nExcel_std_row = nRow + 1;
                    }
                }
            }
            else
            {

                m_nExcel_std_row = 3;

                m_excel_io.set_data(1, "A", "업무일지", Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen);
                m_excel_io.set_data((m_nExcel_std_row - 1), m_cExcel_stdCtxt_column.ToString(), DateTime.Now.ToString("MM월 dd일 ddd"), Microsoft.Office.Interop.Excel.XlRgbColor.rgbDarkGray);
            }
        }

        private struct excel_cell_position
        {
            public int row { get; set; }
            public char column { get; set; }
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
        private System.Timers.Timer m_timer = new System.Timers.Timer(60000);
        private string m_auto_save_file_name = "auto_save_file.json";

        private excel_cell_position m_stCount_cell = new excel_cell_position();
    }
}