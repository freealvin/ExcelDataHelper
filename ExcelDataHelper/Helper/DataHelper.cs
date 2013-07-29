using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ExcelDataHelper.Entity;
using System.Text.RegularExpressions;

namespace ExcelDataHelper.Helper
{
    class DataHelper
    {
        private static Regex RegInvalidDate = new Regex("[0]*", RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace | RegexOptions.Singleline);

        public static bool IsInvalidDate(string strData)
        {
            if (string.IsNullOrEmpty(strData) || RegInvalidDate.Match(strData).Success)
            {
                return true;
            }
            return false;
        }
    }
}
