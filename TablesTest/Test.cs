using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TablesTest
{
    class Test
    {
        static void Main(string[] args)
        {
            string filename;

            Building[] C;
            filename = @"C:\Users\Lubo\Desktop\Centroida\TablesTest\buildingstable.xlsx";
            ExtractBuildings E = new ExtractBuildings(filename);
            C = E.Extract();

            Household[] F;
            filename = @"C:\Users\Lubo\Desktop\Tables\бул. Св.Св. Кирил и Методйй № 24, вх. Г.xls";
            ExtractHouseholds D = new ExtractHouseholds(filename);
            F = D.Extract();

            

            


        }
    }
}
