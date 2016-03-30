using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TablesTest
{
    class ExtractBuildings
    {
        static public string file { set; get; }

        public ExtractBuildings (string a)
        {
            file = a;
        }

        public Building[] Extract()
        {
            Excel.Application app;
            Excel.Workbook wb;
            Excel.Worksheet ws;
            Building[] A;
            int n;

            app = new Excel.Application();
            wb = app.Workbooks.Open(file);
            ws = (Excel.Worksheet)wb.Sheets[1];
            n = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row - 1;
            A = new Building[n];
            for (int i = 0; i < n; i++)
            {
                A[i] = new Building(ws, i + 2);
            }
            return A;

        }
    }
}
