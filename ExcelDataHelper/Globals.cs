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
        /// Ĭ�������ַ���
        /// </summary>
        public static string DefaultConnectionString
        {
            get { return Globals.defaultConnectionString; }
        }

        /// <summary>
        /// ��ŵ������õ��ļ������·��
        /// </summary>
        public static string ConfigXmlPath
        {
            get { return configXmlPath; }
        }

        /// <summary>
        /// ģ���ļ����ļ���
        /// </summary>
        public static string TemplatePath
        {
            get { return Globals.templatePath; }
        }

        /// <summary>
        /// Excel�������ַ������и�{0}��Ҫ�ļ�·���滻
        /// </summary>
        public static string ExcelConnectionString
        {
            get { return Globals.excelConnectionString; }
        }

        /// <summary>
        /// ��ʽ���ַ����ķָ���
        /// </summary>
        public static char FormatStringSplitter
        {
            get { return Globals.formatStringSplitter; }
        }
    }
}
