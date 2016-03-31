using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TablesTest
{
    class Test
    {
        static public string folder = @"C:\Users\Lubo\Desktop\Centroida\Pft";

        static void Main(string[] args)
        {
            string filename;

            string[] temp;

            temp = Directory.GetFiles(folder);
            Household[] A;
            ExtractHouseholds B;
            for (int i = 0; i < temp.Length; ++i)
            {
                B = new ExtractHouseholds(temp[i]);
                A = B.Extract();

            }
        }






    }
}
