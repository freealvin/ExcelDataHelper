using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataHelper.Entity
{
    /// <summary>
    /// 导入数据的信息类，包括一个导入类型的所有信息
    /// </summary>
    public class ImportDataInfo
    {
        private string typeName;
        private string tableName;
        private string connectionString;
        private DestinationType destinationType;
        private bool autoMerge = false;
        private bool oneToMultiRow = false;
        private string multiRowColumnName;
        private bool onlyUpdate = false;

        private List<AttributeInfo> attributeList;
        private List<AttributeInfo> primaryKeyList;

        /// <summary>
        /// 是否只更新
        /// </summary>
        public bool OnlyUpdate
        {
            get { return onlyUpdate; }
            set { onlyUpdate = value; }
        }

        /// <summary>
        /// 该数据的类型名称
        /// </summary>
        public string TypeName
        {
            get { return typeName; }
            set { typeName = value; }
        }

        /// <summary>
        /// 数据库表名
        /// </summary>
        public string TableName
        {
            get { return tableName; }
            set { tableName = value; }
        }

        /// <summary>
        /// 连接字符串
        /// </summary>
        public string ConnectionString
        {
            get { return connectionString; }
            set { connectionString = value; }
        }

        /// <summary>
        /// 数据目标名称
        /// </summary>
        public DestinationType DestinationType
        {
            get { return destinationType; }
            set { destinationType = value; }
        }

        /// <summary>
        /// 是否自动Merge，否的话就是一直插入
        /// </summary>
        public bool AutoMerge
        {
            get { return autoMerge; }
            set { autoMerge = value; }
        }

        /// <summary>
        /// 是否是一行导入为多行
        /// </summary>
        public bool OneToMultiRow
        {
            get { return oneToMultiRow; }
            set { oneToMultiRow = value; }
        }

        /// <summary>
        /// 多行的列名
        /// </summary>
        public string MultiRowColumnName
        {
            get { return multiRowColumnName; }
            set { multiRowColumnName = value; }
        }

        /// <summary>
        /// 导入的属性（列）列表
        /// </summary>
        public List<AttributeInfo> AttributeList
        {
            get { return attributeList; }
            set { attributeList = value; }
        }

        /// <summary>
        /// 主键的列表
        /// </summary>
        public List<AttributeInfo> PrimaryKeyList
        {
            get { return primaryKeyList; }
            set { primaryKeyList = value; }
        }
    }
}
