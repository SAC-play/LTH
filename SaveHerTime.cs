using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Markup;

namespace LTH
{
    public class SaveHerTime
    {
        public bool set_time_unit(double time_unit)
        {
            int gap = (int)(time_unit - m_dTime_unit);
            bool bEndTimeChanged = false;
            DateTime temp_end_time = m_end_time.AddMinutes((double)gap);

            if ((DateTime.Compare(DateTime.Now.AddMinutes(1), temp_end_time) < 0))
            {
                m_end_time = temp_end_time;
                m_dTime_unit = time_unit;
                bEndTimeChanged = true;
            }
            else
            {
                m_dTime_unit = time_unit;
            }

            return bEndTimeChanged;
        }

        public void start_timer()
        {
            m_timer.Elapsed += on_tick;
            m_timer.Start();
        }

        public class SaveHerTimeEventArgs : EventArgs
        {
            public DateTime BeginTime { get; set; }
            public DateTime EndTime { get; set; }
            public DateTime FutureEndTime { get; set; }
        }

        private void on_tick(Object source, ElapsedEventArgs e)
        {
            //string begin_end_time = "begin time : " + m_begin_time.ToString() + ", end time : " + m_end_time.ToString();
            //string elapseTIme_time = "elapse time : " + e.SignalTime.ToString();

            //Console.WriteLine("begin time : " + m_begin_time.ToString());
            //Console.WriteLine("end time : " + m_end_time.ToString());

            //MessageBox.Show("[on_tick]\n" + begin_end_time+"\n"+elapseTIme_time);

            if (e.SignalTime.Hour >= m_end_time.Hour &&
                e.SignalTime.Minute >= m_end_time.Minute)
            {
                if (this.m_elapsed != null)
                {
                    SaveHerTimeEventArgs args = new SaveHerTimeEventArgs();

                    args.BeginTime = m_begin_time;
                    args.EndTime = m_end_time;
                    args.FutureEndTime = m_end_time.AddMinutes(m_dTime_unit);

                    m_elapsed(this, args);
                }

                //modify begin and end time

                if(m_bChangeBeginTime)
                {
                    m_begin_time = m_end_time;
                }

                m_end_time = m_end_time.AddMinutes(m_dTime_unit);

                m_bChangeBeginTime = false;
            }
        }

        public void calculate_start_time()
        {
            DateTime dt = DateTime.Now;

            if(dt.Hour >= 9)
            {
                if(dt.Minute >=0 && dt.Minute < 15)
                {
                    m_begin_time = new DateTime(dt.Year,dt.Month,dt.Day,9,0,0);
                    m_end_time = new DateTime(dt.Year, dt.Month, dt.Day, dt.Hour,0,0);
                    m_end_time = m_end_time.AddMinutes(m_dTime_unit);
                }
                else if(dt.Minute >= 15 && dt.Minute < 30)
                {
                    m_begin_time = new DateTime(dt.Year, dt.Month, dt.Day, 9, 0, 0);
                    m_end_time = new DateTime(dt.Year, dt.Month, dt.Day, dt.Hour,15, 0);
                    m_end_time = m_end_time.AddMinutes(m_dTime_unit);
                }
                else if(dt.Minute >= 30 && dt.Minute < 45 )
                {
                    m_begin_time = new DateTime(dt.Year, dt.Month, dt.Day, 9, 0, 0);
                    m_end_time = new DateTime(dt.Year, dt.Month, dt.Day, dt.Hour, 30, 0);
                    m_end_time = m_end_time.AddMinutes(m_dTime_unit);
                }
                else
                {
                    m_begin_time = new DateTime(dt.Year, dt.Month, dt.Day, 9, 0, 0);
                    m_end_time = new DateTime(dt.Year, dt.Month, dt.Day, dt.Hour, 45, 0);
                    m_end_time = m_end_time.AddMinutes(m_dTime_unit);
                }
            }
            else
            {
                m_begin_time = new DateTime(dt.Year, dt.Month, dt.Day, dt.Hour, 0, 0);
                m_end_time = new DateTime(dt.Year, dt.Month, dt.Day, 9, 0, 0);
                m_end_time = m_end_time.AddMinutes(m_dTime_unit);
            }
        }

        public DateTime find_monday_date()
        {
            DateTime dt_now = DateTime.Now;
            DateTime monday_dt;

            if(dt_now.DayOfWeek == DayOfWeek.Monday)
            {
                monday_dt = dt_now;
            }
            else if(dt_now.DayOfWeek == DayOfWeek.Tuesday)
            {
                monday_dt = dt_now.AddDays(-1);
            }
            else if (dt_now.DayOfWeek == DayOfWeek.Wednesday)
            {
                monday_dt = dt_now.AddDays(-2);
            }
            else if (dt_now.DayOfWeek == DayOfWeek.Thursday)
            {
                monday_dt = dt_now.AddDays(-3);
            }
            else if (dt_now.DayOfWeek == DayOfWeek.Friday)
            {
                monday_dt = dt_now.AddDays(-4);
            }
            else if (dt_now.DayOfWeek == DayOfWeek.Saturday)
            {
                monday_dt = dt_now.AddDays(-5);
            }
            else //Sunday
            {
                monday_dt = dt_now.AddDays(-6);
            }

            return monday_dt;
        }

        public void stop_timer()
        {
            m_timer.Stop();
        }

        public void destroy_timer()
        {
            m_timer.Dispose();
        }

        public DateTime BeginTime
        {
            get => m_begin_time;
            set => m_begin_time = value;
        }

        public DateTime EndTime
        {
            get => m_end_time;
            set => m_end_time = value;
        }

        public double TimeUnit
        {
            get => m_dTime_unit;
        }

        public void add_elapsed_handler(EventHandler ev)
        {
            m_elapsed += ev;
        }

        public bool ChangeBeginTIme
        {
            set => m_bChangeBeginTime = value;
        }

        double m_dTime_unit = 0;
        System.Timers.Timer m_timer = new System.Timers.Timer(60000); //1 min
        private DateTime m_begin_time;
        private DateTime m_end_time;
        private event EventHandler m_elapsed;
        private bool m_bChangeBeginTime = false;
    }
}