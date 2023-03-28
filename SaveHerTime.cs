using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace LTH
{
    class SaveHerTime
    {
        public void set_time_unit(double time_unit)
        {
            int gap = (int)(time_unit - m_dTime_unit);

            if(gap >= 0)
            {
                m_end_time = m_end_time.AddMinutes((double)gap);
                m_dTime_unit = time_unit;
            }
            else
            {
                m_dTime_unit = time_unit;
            }
        }

        public void start_timer()
        {
            m_timer.Elapsed += on_tick;
            m_timer.Start();
        }

        private void on_tick(Object source, ElapsedEventArgs e)
        {
            //Console.WriteLine("end time : " + m_end_time.ToString());

            if (e.SignalTime.Hour == m_end_time.Hour &&
                e.SignalTime.Minute == m_end_time.Minute)
            {
                //Console.WriteLine("Elapsed~~");

                m_begin_time  = e.SignalTime;
                m_end_time = e.SignalTime.AddMinutes(m_dTime_unit);

                m_timer.Start();
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

        double m_dTime_unit = 0;
        System.Timers.Timer m_timer = new System.Timers.Timer(10000); //1 min
        private DateTime m_begin_time;
        private DateTime m_end_time;
    }
}