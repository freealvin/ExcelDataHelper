using System;
using System.Collections.Generic;
using System.Text;
using System.Web;

namespace DataImportLib.Helper
{
    public class FileHelper
    {
        /// <summary>
        /// 得到文件的扩展名，帮助方法
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string GetExtension(string fileName)
        {
            string extension = "";
            int pos = fileName.LastIndexOf('.');
            if (pos >= 0)
            {
                extension = fileName.Substring(pos, fileName.Length - pos);
            }

            return extension;
        }

        /// <summary>
        /// 下载文件的代码，支持断点续传
        /// </summary>
        /// <param name="docUrl"></param>
        /// <param name="Request"></param>
        /// <param name="Response"></param>
        public static void DownloadDoc(string docUrl, System.Web.HttpRequest Request, System.Web.HttpResponse Response)
        {
            //得到文件句柄
            System.IO.FileInfo thefile = new System.IO.FileInfo(docUrl);

            //创建文件流实例
            System.IO.Stream iStream = null;

            // Buffer to read 10K bytes in chunk:

            byte[] buffer = new Byte[10240];

            // Length of the file:
            int length;

            // Total bytes to read:
            long dataToRead;
            try
            {
                iStream = new System.IO.FileStream(docUrl,
                                                    System.IO.FileMode.Open,
                                                    System.IO.FileAccess.Read,
                                                    System.IO.FileShare.Read);

                dataToRead = iStream.Length;

                Response.Clear();

                long p = 0;
                Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(docUrl, System.Text.Encoding.UTF8));
                if (Request.Headers["Range"] != null)
                {
                    Response.StatusCode = 206;
                    p = long.Parse(Request.Headers["Range"].Replace("bytes=", "").Replace("-", ""));
                }
                if (p != 0)
                {
                    Response.AddHeader("Content-Range", "bytes " + p.ToString() + "-" + ((long)(dataToRead - 1)).ToString() + "/" + dataToRead.ToString());
                }
                Response.ContentType = "application/octet-stream";
                Response.AddHeader("Content-Length", ((long)(dataToRead - p)).ToString());

                iStream.Position = p;
                dataToRead = dataToRead - p;
                while (dataToRead > 0)
                {
                    // Verify that the client is connected.
                    if (Response.IsClientConnected)
                    {
                        // Read the data in buffer.
                        length = iStream.Read(buffer, 0, 10240);

                        // Write the data to the current output stream.
                        Response.OutputStream.Write(buffer, 0, length);

                        // 向页面输出数据
                        Response.Flush();

                        buffer = new Byte[10240];
                        dataToRead = dataToRead - length;
                    }
                    else
                    {
                        //prevent infinite loop if user disconnects
                        dataToRead = -1;
                    }
                }

            }
            catch (Exception ex)
            {
                // Trap the error, if any.
                Response.Write("Error : " + ex.Message);
            }
            finally
            {
                if (iStream != null)
                {
                    //Close the file.
                    iStream.Close();
                }
                Response.End();
            }
        }

    }
}
