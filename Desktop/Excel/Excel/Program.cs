using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Windows;

namespace Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            _Application excel = new _Excel.Application();
            Workbook wb;
            Worksheet ws;
            int i = 1, j = 1;
            wb = excel.Workbooks.Open(@"C:\Users\Семен\Desktop\Test");
            ws = wb.Worksheets[1];
            if (ws.Cells[i, j].Value2 != null)
               MessageBox.Show(ws.Cells[i, j].Value2);
            else
              MessageBox.Show("");

        }
    }
}
