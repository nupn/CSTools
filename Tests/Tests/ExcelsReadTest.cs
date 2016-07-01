using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data.OleDb;

namespace Tests
{

    class ExcelsReadTest
    {
        
        static void ReadExcelDate()
        {
            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                excelApp = new Excel.Application();
                String path = System.IO.Path.GetFullPath("../../test.xlsx");
                Console.WriteLine(path);
                wb = excelApp.Workbooks.Open(path);

                ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);
        
                //Excel.Range rng = ws.UsedRange;
                //Excel.Range rng = ws.Range[ws.Cells[2,1], ws.Cells[5,3]];
                Excel.Range rng = ws.get_Range(ws.Cells[1, 1], ws.Cells[3, 2]);

                System.Array result = (System.Array)rng.Value;
                Console.WriteLine(rng.Value.GetType());

                foreach (object cols in result)
                {
                    if (cols != null)
                    {
                        String str = cols.ToString();
                        Console.WriteLine(str);
                    }
                    else
                    {
                        Console.WriteLine("null");
                    }
                    /*
                    foreach (String eachItem in cols)
                        Console.WriteLine(eachItem);*/
                }

                Console.WriteLine(rng.Value.GetType());
                System.Array arr = (System.Array)rng.Value;

                Console.WriteLine(String.Format("{0} {1}", arr.GetLength(0), arr.GetLength(1)));

                for (int r = 1; r <= arr.GetLength(0); r++)
                {
                    for (int c = 1; c <= arr.GetLength(1); c++)
                    {
                        if (arr.GetValue(r, c) != null)
                        {
                            Console.WriteLine(arr.GetValue(r, c).ToString());
                        }
                    }
                }

                Excel.Range range2 = ws.get_Range(ws.Cells[1, 1], ws.Cells[1, 2]);
                Console.WriteLine(range2.Value.GetHashCode());

                range2 = ws.get_Range(ws.Cells[2, 1], ws.Cells[2, 2]);
                Console.WriteLine(range2.Value.GetHashCode());

                range2 = ws.get_Range(ws.Cells[3, 1], ws.Cells[3, 2]);
                Console.WriteLine(range2.Value.GetHashCode());


                wb.Close(true);
                excelApp.Quit();

            }
            finally
            {
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(excelApp);
            }

        }

        private static void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (System.Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Hello world");   
            ReadExcelDate();
            Console.ReadKey();
        }
    }
}
