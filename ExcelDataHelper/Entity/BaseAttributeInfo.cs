using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataHelper.Entity
{
    /// <summary>
    /// 一列数据属性的基类
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
        /// 数据库表列名
        /// </summary>
        public string ColumnName
        {
            get { return this.columnName; }
            set { this.columnName = value; }
        }

        /// <summary>
        /// 列的数据库类型，默认为字符串
        /// </summary>
        public string ValueType
        {
            get { return valueType; }
            set { valueType = value; }
        }
        /// <summary>
        /// 确定的值，如果有外连接的表则是确定的键的值
        /// </summary>
        public string DetermineValue
        {
            get { return determineValue; }
            set { determineValue = value; }
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
        public string[] FormatStrings
        {
            get { return this.formatStrings; }
            set { this.formatStrings = value; }
        }

        /// <summary>
        /// 是否需要加密
        /// </summary>
        public bool NeedEncrypt
        {
            get { return needEncrypt; }
            set { needEncrypt = value; }
        }
    }
}
