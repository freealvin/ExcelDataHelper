using System;
using System.IO;
using System.Text;
using System.Data;
using System.Reflection;
using System.Diagnostics;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;

//using cfg = System.Configuration;
namespace ExcelDataHelper.Helper
{
    public class ExcelHelper
    {
        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="eDataTable"></param>
        /// <param name="filePath"></param>
        public static void ExportExcel(DataTable eDataTable, string filePath)
        {
            //写在外面是为了完成后释放
            Excel.ApplicationClass ExcelApp = null;
            Excel.Workbook ExcelBook = null;
            Excel.Worksheet ExcelSheet = null;
            Excel.Range range = null;
            try
            {
                ExcelApp = new Excel.ApplicationClass(); 
                ExcelApp.Visible = false;
                ExcelBook = ExcelApp.Workbooks.Add(1);
                ExcelSheet = (Excel.Worksheet)ExcelBook.Worksheets[1];

                int colCount = eDataTable.Columns.Count;
                int rowCount = eDataTable.Rows.Count;
                object[,] dataArray = new object[rowCount + 1, colCount];

                ExcelApp.Visible = false;

                //写列标题
                for (int i = 0; i < colCount; i++)
                {
                    dataArray[0,i] = eDataTable.Columns[i].ColumnName;
                }

                //写值
                for (int r = 0; r < rowCount; r++)
                {
                    for (int i = 0; i < colCount; i++)
                    {
                        //ExcelSheet.Cells[r + 2, i + 1] = eDataTable.Rows[r][i];
                        dataArray[r + 1, i] = eDataTable.Rows[r][i];
                    }
                }
                //整体写入Excel
                range = ExcelSheet.get_Range("A1", ExcelSheet.Cells[rowCount + 1, colCount]);
                range.NumberFormat = "@";
                range.Value2 = dataArray;

                ExcelBook.Saved = true;
                object missing = Missing.Value;
                ExcelBook.SaveCopyAs(filePath);
                ExcelBook.Close(false, missing, missing);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //写这么多是为了正常关闭Excel进程，要不可能关不掉
                if (range != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    range = null;
                }
                if (ExcelSheet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelSheet);
                    ExcelSheet = null;
                }
                if (ExcelBook != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelBook);
                    ExcelBook = null;
                }
                if (ExcelApp != null)
                {
                    ExcelApp.Workbooks.Close();
                    ExcelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
                    ExcelApp = null;
                }
                GC.Collect();
            }
        }
    }

}