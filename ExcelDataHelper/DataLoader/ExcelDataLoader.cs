using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace ExcelDataHelper.DataLoader
{
    public class ExcelDataLoader
    {
        //private static string ExcelConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source ={0};Extended Properties = Excel 8.0;IMEX=1";

        /// <summary>
        /// ��excel�ļ��еõ�DataTable
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public DataTable getDataTableFromExcel(string filePath)
        {
            //����Excel�ļ��������ַ���
            string connString = string.Format(Globals.ExcelConnectionString, filePath);
            OleDbConnection oleConn = new OleDbConnection(connString);
            OleDbDataAdapter oleDa = new OleDbDataAdapter("SELECT * FROM [sheet1$]", oleConn);
            DataTable dt = new DataTable();
            try
            {
                oleDa.Fill(dt);
            }
            catch (OleDbException ex)
            {
                throw ex;
            }
            finally
            {
                if (oleConn.State != ConnectionState.Closed)
                {
                    oleConn.Close();
                }
            }
            return dt;
        }
    }
}
