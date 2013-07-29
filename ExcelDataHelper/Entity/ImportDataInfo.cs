using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataHelper.Entity
{
    /// <summary>
    /// �������ݵ���Ϣ�࣬����һ���������͵�������Ϣ
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
        /// �Ƿ�ֻ����
        /// </summary>
        public bool OnlyUpdate
        {
            get { return onlyUpdate; }
            set { onlyUpdate = value; }
        }

        /// <summary>
        /// �����ݵ���������
        /// </summary>
        public string TypeName
        {
            get { return typeName; }
            set { typeName = value; }
        }

        /// <summary>
        /// ���ݿ����
        /// </summary>
        public string TableName
        {
            get { return tableName; }
            set { tableName = value; }
        }

        /// <summary>
        /// �����ַ���
        /// </summary>
        public string ConnectionString
        {
            get { return connectionString; }
            set { connectionString = value; }
        }

        /// <summary>
        /// ����Ŀ������
        /// </summary>
        public DestinationType DestinationType
        {
            get { return destinationType; }
            set { destinationType = value; }
        }

        /// <summary>
        /// �Ƿ��Զ�Merge����Ļ�����һֱ����
        /// </summary>
        public bool AutoMerge
        {
            get { return autoMerge; }
            set { autoMerge = value; }
        }

        /// <summary>
        /// �Ƿ���һ�е���Ϊ����
        /// </summary>
        public bool OneToMultiRow
        {
            get { return oneToMultiRow; }
            set { oneToMultiRow = value; }
        }

        /// <summary>
        /// ���е�����
        /// </summary>
        public string MultiRowColumnName
        {
            get { return multiRowColumnName; }
            set { multiRowColumnName = value; }
        }

        /// <summary>
        /// ��������ԣ��У��б�
        /// </summary>
        public List<AttributeInfo> AttributeList
        {
            get { return attributeList; }
            set { attributeList = value; }
        }

        /// <summary>
        /// �������б�
        /// </summary>
        public List<AttributeInfo> PrimaryKeyList
        {
            get { return primaryKeyList; }
            set { primaryKeyList = value; }
        }
    }
}
