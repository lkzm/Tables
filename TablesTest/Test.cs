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
            

            string[] temp;

            temp = Directory.GetFiles(folder);
            Building[] A = new Building[temp.Length];
            ExtractHouseholds B;
            Excel.Application app = new Excel.Application();
            B = new ExtractHouseholds(app);
            for (int i = 0; i < temp.Length; ++i)
            {
                B.SetFile(temp[i]);
                A[i] = new Building (B.Extract());

            }
        }






    }
}
