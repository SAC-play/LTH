using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace LTH
{
    class Excel_io
    {
        public struct excel_io_data
        {
            public int nRow;
            public string strColumn;
            public string strData;
        }

        public string create_file(string file_path_and_name)
        {
            Application excelApp = new Application();

            if (excelApp == null )
            {
                return "Application is null";
            }

            // Make the object visible.
            // excelApp.Visible = true;

            if (System.IO.File.Exists(AppDomain.CurrentDomain.BaseDirectory.ToString() + file_path_and_name))
            {
                m_file_name = AppDomain.CurrentDomain.BaseDirectory.ToString() + file_path_and_name;

                excelApp.Quit();

                Marshal.ReleaseComObject(excelApp);

                return file_path_and_name + " is existed file";
            }

            var wkbooks = excelApp.Workbooks;
            var wb = wkbooks.Add();


            // !!caution!! microsoft interop should make object below if true.
            // if not, interop remains in backround process
#if true
            var worksheets = wb.Worksheets;
            _Worksheet ws = worksheets[1];
#else
            _Worksheet workSheet = wb.Worksheets[1];
#endif

            ws.Name = "업무일지";

            m_file_name = AppDomain.CurrentDomain.BaseDirectory.ToString() + file_path_and_name;

            wb.SaveAs(m_file_name);

            Marshal.ReleaseComObject(ws);
            Marshal.ReleaseComObject(worksheets);

            wb.Close();
            Marshal.ReleaseComObject(wb);

            wkbooks.Close();
            Marshal.ReleaseComObject(wkbooks);

            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);

            return "success";
        }

        public string get_file_name()
        {
            return m_file_name;
        }

        public void set_data(int row, string column, string data)
        {
            var dict_key = (row.ToString() + column);

            if (m_dict_data.ContainsKey(dict_key))
            {
                excel_io_data stData;
                stData.nRow = row;
                stData.strColumn = column;

                string total_str = m_dict_data[dict_key].strData + "\n" + data;

                stData.strData = total_str;

                m_dict_data[dict_key] = stData;
            }
            else
            {
                excel_io_data stData;
                stData.nRow = row;
                stData.strColumn = column;
                stData.strData = data;

                m_dict_data.Add(dict_key, stData);
            }
        }

        public bool sync_data()
        {
            Application excelApp = new Application();

            if (excelApp == null)
            {
                return false;
            }

            if (!System.IO.File.Exists(m_file_name))
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);

                return false;
            }

            var workbooks = excelApp.Workbooks;
            var wb = workbooks.Open(m_file_name);

            var worksheets = wb.Worksheets;
            _Worksheet ws = worksheets[1];

            if(ws == null)
            {
                wb.Close();
                excelApp.Quit();

                Marshal.ReleaseComObject(worksheets);
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(workbooks);
                Marshal.ReleaseComObject(excelApp);

                return false;
            }

            // add data into cells
            {
                var key_list = m_dict_data.Keys.ToList();
                key_list.Sort();

                var cells = ws.Cells;

                foreach (var key_item in key_list)
                {
                    excel_io_data stData = m_dict_data[key_item];

                    cells[stData.nRow, stData.strColumn] = stData.strData;
                }

                Marshal.ReleaseComObject(cells);
            }


            // to autofit column
            {
                var key_list = m_dict_data.Keys.ToList();
                key_list.Sort();

                List<int> column_num_list = new List<int>();
                int current_pivot = 0;

                foreach (var key_item in key_list)
                {
                    excel_io_data stData = m_dict_data[key_item];

                    int column_num = ((int)char.Parse(stData.strColumn) - 64);

                    if (column_num > current_pivot)
                    {
                        column_num_list.Add(column_num);
                        current_pivot = column_num;
                    }
                }

                var columns = ws.Columns;

                foreach (var nColumn in column_num_list)
                {
                    var cl = columns[nColumn];
                    cl.Autofit();
                    Marshal.ReleaseComObject(cl);
                }

                Marshal.ReleaseComObject(columns);
            }

            wb.Save();

            Marshal.ReleaseComObject(ws);
            ws = null;
            Marshal.ReleaseComObject(worksheets);
            worksheets = null;

            wb.Close();
            Marshal.ReleaseComObject(wb);

            workbooks.Close();
            Marshal.ReleaseComObject(workbooks);

            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);

            return true;
        }

        public Dictionary<string, excel_io_data> get_data()
        {
            return m_dict_data;
        }

        private Dictionary<string, excel_io_data> m_dict_data = new Dictionary<string, excel_io_data>();
        private string m_file_name = "";
    }
}
