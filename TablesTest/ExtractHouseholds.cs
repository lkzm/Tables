using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TablesTest
{
    class ExtractHouseholds
    {
        static public string file { set; get; }

        public ExtractHouseholds(string a)
        {
            file = a;
        }
        public int Wsn (Excel.Workbook wb)
        {
            Excel.Worksheet ws;
            int i = 1;
            ws = (Excel.Worksheet)wb.Sheets[i];

            while (ws.Cells.Find("С  П  Р  А  В  К  А", Type.Missing,
Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole,
Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
Type.Missing, Type.Missing) == null)
            {

                i++;
                ws = (Excel.Worksheet)wb.Sheets[i];

            }
            return i;
        }
        public Household[] Extract()
        {
            Excel.Application app;
            Excel.Workbook wb;
            Excel.Worksheet ws;
            Household[] A;
            int r, n, i = 1;

            app = new Excel.Application();
            wb = app.Workbooks.Open(file);
            ws = (Excel.Worksheet)wb.Sheets[Wsn(wb)];
            n = 0;
            Excel.Range range = ws.get_Range("A" + i);



            while (!Int32.TryParse(range.Text.ToString(), out r))
            {
                ++i;
                range = ws.get_Range("A" + i);
            }
            n = i;
            while (Int32.TryParse(range.Text.ToString(), out r))
            {
                ++i;
                range = ws.get_Range("A" + i);

            }
            A = new Household[i-n];
            

            for (int j = n; j<i; ++j)
            {
                A[j - n] = new Household(ws, j);
            }



            return A;

        }
    }
}
