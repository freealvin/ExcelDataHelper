using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataHelper.Entity
{
    /// <summary>
    /// һ���������ԵĻ���
    /// </summary>
    public class BaseAttributeInfo
    {
        private string columnName;
        private string valueType;
        private string determineValue;
        private bool isPrimaryKey = false;
        private string outerTable;
        private string keyColumn;
        private string valueColumn;
        private string[] formatStrings;
        private bool needEncrypt;

        /// <summary>
        /// ���ݿ������
        /// </summary>
        public string ColumnName
        {
            get { return this.columnName; }
            set { this.columnName = value; }
        }

        /// <summary>
        /// �е����ݿ����ͣ�Ĭ��Ϊ�ַ���
        /// </summary>
        public string ValueType
        {
            get { return valueType; }
            set { valueType = value; }
        }
        /// <summary>
        /// ȷ����ֵ������������ӵı�����ȷ���ļ���ֵ
        /// </summary>
        public string DetermineValue
        {
            get { return determineValue; }
            set { determineValue = value; }
        }
        /// <summary>
        /// �Ƿ���������Ĭ��false
        /// </summary>
        public bool IsPrimaryKey
        {
            get { return this.isPrimaryKey; }
            set { this.isPrimaryKey = value; }
        }
        /// <summary>
        /// ��Ҫ�����ӵı�
        /// </summary>
        public string OuterTable
        {
            get { return this.outerTable; }
            set { this.outerTable = value; }
        }
        /// <summary>
        /// �������������
        /// </summary>
        public string KeyColumn
        {
            get { return this.keyColumn; }
            set { this.keyColumn = value; }
        }

        /// <summary>
        /// ������ֵ������
        /// </summary>
        public string ValueColumn
        {
            get { return this.valueColumn; }
            set { this.valueColumn = value; }
        }

        /// <summary>
        /// ��ʽ���ַ���
        /// </summary>
        public string[] FormatStrings
        {
            get { return this.formatStrings; }
            set { this.formatStrings = value; }
        }

        /// <summary>
        /// �Ƿ���Ҫ����
        /// </summary>
        public bool NeedEncrypt
        {
            get { return needEncrypt; }
            set { needEncrypt = value; }
        }
    }
}
