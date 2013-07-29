using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;

namespace DataImportLib
{
    public class Globals
    {
        private static string excelConnectionString = ConfigurationManager.ConnectionStrings["excelConnectionString"].ConnectionString;
        private static string defaultConnectionString = ConfigurationManager.ConnectionStrings["defaultConnectionString"].ConnectionString;
        private static string configXmlPath = ConfigurationManager.AppSettings["configXmlPath"];
        private static string templatePath = ConfigurationManager.AppSettings["templatePath"];
        private static char formatStringSplitter = char.Parse(ConfigurationManager.AppSettings["formatStringSplitter"]);

        /// <summary>
        /// 默认连接字符串
        /// </summary>
        public static string DefaultConnectionString
        {
            get { return Globals.defaultConnectionString; }
        }

        /// <summary>
        /// 存放导入配置的文件夹相对路径
        /// </summary>
        public static string ConfigXmlPath
        {
            get { return configXmlPath; }
        }

        /// <summary>
        /// 模板文件的文件夹
        /// </summary>
        public static string TemplatePath
        {
            get { return Globals.templatePath; }
        }

        /// <summary>
        /// Excel的连接字符串，有个{0}需要文件路径替换
        /// </summary>
        public static string ExcelConnectionString
        {
            get { return Globals.excelConnectionString; }
        }

        /// <summary>
        /// 格式化字符串的分隔符
        /// </summary>
        public static char FormatStringSplitter
        {
            get { return Globals.formatStringSplitter; }
        }
    }
}
