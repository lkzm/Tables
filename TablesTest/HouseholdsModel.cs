using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace TablesTest
{
    class Household
    {
        public int Number { set; get; }
        public String Name { set; get; }
        static public int Inhab { set; get; }

        public Household(Excel.Worksheet ws, int i)
        {
            int p;

            Excel.Range range1, range2, range3;
            range1 = ws.get_Range("A" + i);
            range2 = ws.get_Range("B" + i);
            range3 = ws.get_Range("C" + i);

            Number = Convert.ToInt32(range1.Text.ToString());
            Name = range2.Text.ToString();
            if (range3.Text.ToString() == "")
            {
                Inhab = -1;
            }
            else
            {
                if (Int32.TryParse(range3.Text.ToString(), out p))
                {

                    Inhab = p;
                }
                else
                {
                    Inhab = 0;
                }
            }

            




        }
    }
}
