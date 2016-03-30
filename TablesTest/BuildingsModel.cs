using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TablesTest
{
    class Building
    {
        public string Address { set; get; }
        public string Location { set; get; }
        public int FloorsNumber { set; get; }

        public Building (Excel.Worksheet ws, int i)
        {
            Excel.Range range1, range2, range3;
            range1 = ws.get_Range("A" + i);
            range2 = ws.get_Range("B" + i);
            range3 = ws.get_Range("C" + i);


            Address = range1.Text.ToString();
            Location = range2.Text.ToString();
            String asd = range3.Text.ToString();
            if (range3.Text.ToString() == "" || range3.Text.ToString () == "0+5")   
            {
                FloorsNumber = 1;
            }
            else
            {
                FloorsNumber = Convert.ToInt32(range3.Text.ToString());
            }
            
        }

    }
}
