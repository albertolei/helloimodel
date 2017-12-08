using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace NImodel.Utils
{
    class ExcelUtil
    {
        public static void writeToExcel(Dictionary<string, Dictionary<string, string>> imodelInfo, string filename)
        {
            Application excelApp = new ApplicationClass();
            Workbook workBook;
            if (excelApp == null)
            {
                Console.WriteLine("There's no excel exits.");
            }
            else
            {
                if (File.Exists(filename))
                {
                    workBook = excelApp.Workbooks.Open(filename, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);  
                }
                else
                {
                    workBook = excelApp.Workbooks.Add(true);
                    for (int i = 1; i < imodelInfo.Keys.Count; i++)
                    {
                        workBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }
                    int k = 1;
                    foreach (string key in imodelInfo.Keys)
                    {
                        Worksheet workSheet;
                        try 
                        {
                            workSheet = (Worksheet)workBook.Worksheets[k++];
                            workSheet.Name = key;
                            workSheet.Activate();
                            Dictionary<string, string> elementsInfo = imodelInfo[key];
                            int i = 1;
                            foreach (string elementsInfoKey in elementsInfo.Keys)
                            {
                                workSheet.Cells[1, i] = elementsInfoKey;
                                workSheet.Cells[2, i] = elementsInfo[elementsInfoKey];
                                i++;
                            }
                            workSheet.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workSheet.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                            workSheet.Columns.AutoFit();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            Console.WriteLine(ex.StackTrace);
                        }
                    }
                }
                workBook.SaveAs(filename);
                workBook.Close();
            }
            excelApp.Quit();
            GC.Collect();
        }

        public static void testExcel(string path)
        {
            Application excelApp = new ApplicationClass();
            Workbook workBook;
            if (excelApp == null)
            {
                Console.WriteLine("There's no excel exits.");
            }
            else
            {
                if (File.Exists(path))
                {
                    workBook = excelApp.Workbooks.Open(path, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);
                }
                else
                {
                    workBook = excelApp.Workbooks.Add(true);
                }
                for (int i = 1; i <= 4; i++)
                {
                    workBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
                for (int i = 1; i <= 5; i++)
                {
                    //Console.WriteLine("i: " + i);
                    Worksheet workSheet = workBook.Worksheets[i] as Worksheet;
                    workSheet.Activate();
                    workSheet.Cells[1, 3] = "(1,3)Content";
                }
                workBook.SaveAs(path);
                workBook.Close();

            }
            excelApp.Quit();
            GC.Collect();
        }
    }
}
