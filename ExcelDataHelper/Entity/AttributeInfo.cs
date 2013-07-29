using System;
using System.Collections.Generic;
using System.Text;

namespace DataImportLib.Entity
{
    /// <summary>
    /// 导入列的信息类，包括一个列的所有信息
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
        /// 导入数据源的列名
        /// </summary>
        public string SourceName
        {
            get { return this.sourceName; }
            set { this.sourceName = value; }
        }

        /// <summary>
        /// 是否将前后的空格去掉,默认为false
        /// </summary>
        public bool Trim
        {
            get { return this.trim; }
            set { this.trim = value; }
        }

        /// <summary>
        /// 是否是确定的几个值中的一个
        /// </summary>
        public bool IsCertainValue
        {
            get { return this.isCertainValue; }
            set { this.isCertainValue = value; }
        }

        /// <summary>
        /// 示例数据
        /// </summary>
        public string ExampleData
        {
            get { return exampleData; }
            set { exampleData = value; }
        }

        /// <summary>
        /// 确定的值列表
        /// </summary>
        public List<CertainValueInfo> CertainValueList
        {
            get { return certainValueList; }
            set { certainValueList = value; }
        }

        /// <summary>
        /// 表示此列可以确定的其他的列的列表
        /// </summary>
        public List<DetermineColumnInfo> DetermineColumnList
        {
            get { return determineColumnList; }
            set { determineColumnList = value; }
        }

    }
}
#region 废代码
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
        /// 数据库表列名
        /// </summary>
        public string ColumnName
        {
            get { return this.columnName; }
            set { this.columnName = value; }
        }

        /// <summary>
        /// 值的数据库类型,默认为字符串类型
        /// </summary>
        public string ValueType
        {
            get { return this.valueType; }
            set { this.valueType = value; }
        }

        /// <summary>
        /// 是否是主键，默认false
        /// </summary>
        public bool IsPrimaryKey
        {
            get { return this.isPrimaryKey; }
            set { this.isPrimaryKey = value; }
        }

        /// <summary>
        /// 需要外联接的表
        /// </summary>
        public string OuterTable
        {
            get { return this.outerTable; }
            set { this.outerTable = value; }
        }

        /// <summary>
        /// 外联表键的列名
        /// </summary>
        public string KeyColumn
        {
            get { return this.keyColumn; }
            set { this.keyColumn = value; }
        }

        /// <summary>
        /// 外联表值的列名
        /// </summary>
        public string ValueColumn
        {
            get { return this.valueColumn; }
            set { this.valueColumn = value; }
        }

        /// <summary>
        /// 格式化字符串
        /// </summary>
        public string FormatString
        {
            get { return this.formatString; }
            set { this.formatString = value; }
        }*/

#endregion
