using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    class Program
    {
        static void Main(string[] args)
        {
            //create the Application object we can use in the member functions.
            Microsoft.Office.Interop.Excel.Application _excelApp = new Microsoft.Office.Interop.Excel.Application();
            _excelApp.Visible = true;
            //string fileName = "C:\\sampleExcelFile.xlsx";
            string fileName = Path.GetDirectoryName(System.AppDomain.CurrentDomain.BaseDirectory) + "\\DB35_STDA110U.xlsx";
            //open the workbook
            Workbook workbook = _excelApp.Workbooks.Open(fileName,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            for (int ws = 2; ws < workbook.Worksheets.Count; ws++)
            {
                //select the first sheet        
                Worksheet worksheet = (Worksheet)workbook.Worksheets[ws];

                //find the used range in worksheet
                Range excelRange = worksheet.UsedRange;

                //get an object array of all of the cells in the worksheet (their values)
                object[,] valueArray = (object[,])excelRange.get_Value(
                            XlRangeValueDataType.xlRangeValueDefault);

                //access the cells
                for (int row = 1; row <= worksheet.UsedRange.Rows.Count; ++row)
                {
                    for (int col = 1; col <= worksheet.UsedRange.Columns.Count; ++col)
                    {
                        //access each cell
                        //Debug.Print(valueArray[row, col].ToString());
                        var cellValue = valueArray[row, col];
                    }
                }
            }

            //clean up stuffs
            workbook.Close(false, Type.Missing, Type.Missing);
            //Marshal.ReleaseComObject(workbook);

            _excelApp.Quit();
            //Marshal.FinalReleaseComObject(_excelApp);

        }

    }
}
