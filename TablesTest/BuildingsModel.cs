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
        public Household[] households; 

        public Building (Household[] a)
        {
            households = a;
        }

    }
}
