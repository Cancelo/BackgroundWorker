using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace BgWorkerApp
{
    class FileExcel
    {
        static string name = "";
        static string part = "";

        public static string Read(string filePath, int row, int col, DateTime startDate, DateTime endDate)
        {
            string matrixValue = "";

            if (!File.Exists(filePath))
            {
                matrixValue = "noFile";
            }

            var misValue = Type.Missing;
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath, misValue, misValue,
                                        misValue, misValue, misValue, misValue, misValue, misValue,
                                        misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            Excel.Range range = xlWorkSheet.UsedRange;

            int totalRows = range.Rows.Count;

            if ((range.Cells[row, col] as Excel.Range).Value != null)   // !=null almacena name
            {
                name = (range.Cells[row, col] as Excel.Range).Value.ToString().Replace(" ", "");
            }

            if ((range.Cells[row + 6, col] as Excel.Range).Value != null)   // !=null almacena part
            {
                part = (range.Cells[row + 6, col] as Excel.Range).Value.ToString();
            }

            int c = 0; // Contador
            int c1 = 0; // Contador matchDate
            for (int startRow = row + 7; startRow <= 539; startRow++)
            {
                if ((range.Cells[startRow, col] as Excel.Range).Value != null)
                {
                    c++;
                    try
                    {
                        string match = (range.Cells[startRow, col] as Excel.Range).Value.ToString();
                        DateTime matchDate = Convert.ToDateTime(match);

                        // Comparación de fechas
                        if (matchDate >= startDate && matchDate <= endDate)
                        {
                            c1++;
                        }
                    }
                    catch (Exception ex) { }
                }
            }
            matrixValue = name + "^" + part + "^" + c1 + "^" + c;

            // Cerrar
            xlWorkBook.Close(false, misValue, misValue);
            xlApp.Quit();

            // Liberar
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            return matrixValue;
        }

        public static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to release the object(object:{0})", obj.ToString());
            }
            finally
            {
                obj = null;
                GC.Collect();
            }
        }

        public static void CreateReport(List<Course> listCourses)
        {
            foreach(Course s in listCourses)
            {                
                if(Convert.ToInt32(s.DateCount) != 0)
                {
                    Console.WriteLine(s.Name + " [" + s.Part + "]:   " + s.DateCount);
                }                
            }           
        }        
    }
}
