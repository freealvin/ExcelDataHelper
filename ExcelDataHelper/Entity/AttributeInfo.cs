using System;
using System.Collections.Generic;
using System.Text;

namespace DataImportLib.Entity
{
    /// <summary>
    /// �����е���Ϣ�࣬����һ���е�������Ϣ
    /// </summary>
    public class AttributeInfo : BaseAttributeInfo
    {
        private string sourceName;
        private bool trim = false;
        private bool isCertainValue = false;
        private string  exampleData;

        private List<CertainValueInfo> certainValueList;
        private List<DetermineColumnInfo> determineColumnList;
        
        /// <summary>
        /// ��������Դ������
        /// </summary>
        public string SourceName
        {
            get { return this.sourceName; }
            set { this.sourceName = value; }
        }

        /// <summary>
        /// �Ƿ�ǰ��Ŀո�ȥ��,Ĭ��Ϊfalse
        /// </summary>
        public bool Trim
        {
            get { return this.trim; }
            set { this.trim = value; }
        }

        /// <summary>
        /// �Ƿ���ȷ���ļ���ֵ�е�һ��
        /// </summary>
        public bool IsCertainValue
        {
            get { return this.isCertainValue; }
            set { this.isCertainValue = value; }
        }

        /// <summary>
        /// ʾ������
        /// </summary>
        public string ExampleData
        {
            get { return exampleData; }
            set { exampleData = value; }
        }

        /// <summary>
        /// ȷ����ֵ�б�
        /// </summary>
        public List<CertainValueInfo> CertainValueList
        {
            get { return certainValueList; }
            set { certainValueList = value; }
        }

        /// <summary>
        /// ��ʾ���п���ȷ�����������е��б�
        /// </summary>
        public List<DetermineColumnInfo> DetermineColumnList
        {
            get { return determineColumnList; }
            set { determineColumnList = value; }
        }

    }
}
#region �ϴ���
/*
        private string  columnName;
        private string  valueType;
        private bool    isPrimaryKey = false;
        private string  outerTable;
        private string  keyColumn;
        private string  valueColumn;
        private string  formatString;*/
/*
        /// <summary>
        /// ���ݿ������
        /// </summary>
        public string ColumnName
        {
            get { return this.columnName; }
            set { this.columnName = value; }
        }

        /// <summary>
        /// ֵ�����ݿ�����,Ĭ��Ϊ�ַ�������
        /// </summary>
        public string ValueType
        {
            get { return this.valueType; }
            set { this.valueType = value; }
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
        public string FormatString
        {
            get { return this.formatString; }
            set { this.formatString = value; }
        }*/

#endregion
