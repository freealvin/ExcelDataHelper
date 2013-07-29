using System;
using System.Collections.Generic;
using System.Text;
using System.Security.Cryptography;

namespace DataImportLib.Helper
{
    public class EncryptHelper
    {
        /// <summary>
        /// ¼ÓÃÜ×Ö·û´®
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string EncrptString(string str)
        {
            MD5CryptoServiceProvider MD5CSP = new MD5CryptoServiceProvider();
            byte[] MD5Source = System.Text.Encoding.UTF8.GetBytes(str);
            byte[] MD5Out = MD5CSP.ComputeHash(MD5Source);
            return Convert.ToBase64String(MD5Out);
        }
    }
}
