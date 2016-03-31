using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;


namespace TablesTest
{
    class ExtractHouseholds
    {

        static public Excel.Application app { set; get; }
        static public string file;
        public ExtractHouseholds(Excel.Application a) 
        {
            app = a;
        }
        public void SetFile (string a)
        {
            file = a;
        }
        public int Wsn (Excel.Workbook wb)
        {
            Excel.Worksheet ws;
            int i = 1;
            int a;
            ws = (Excel.Worksheet)wb.Sheets[i];
            Excel.Range range;
            bool b = true;
            while (b)
            {
                for (int j = 1; j < 50; ++j)
                {
                    range = ws.get_Range("A" + j);
                    if (Int32.TryParse(range.Text.ToString(), out a)) return i;
                }

                i++;
                ws = (Excel.Worksheet)wb.Sheets[i];

            }
            return i;
        }
        public Household[] Extract()
        {
            
            Excel.Workbook wb = null;
            Excel.Worksheet ws;
            Household[] A = null;
            int r, n, i = 1;
            

            try
            {
                
                wb = wb = app.Workbooks.Open(file);
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
                A = new Household[i - n];


                for (int j = n; j < i; ++j)
                {
                    A[j - n] = new Household(ws, j);
                }
            }
            catch (FileNotFoundException e)
            {
                Console.WriteLine(e);

            }
            finally
            {
                if (wb != null )
                {
                    wb.Close(false);
                }
                

            }


            return A;

        }
    }
}
