using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using DataImportLib.Entity;
using System.Data;
using DataImportLib.DataManager;

namespace DataImportLib
{
    public class ImportHandler
    {
        #region Singletonģʽ
        private static ImportHandler instance = new ImportHandler();

        public static ImportHandler GetInstance()
        {
            return instance;
        }
        #endregion

        /// <summary>
        /// �������ݵĹ�ϣ����Ϊ�������ƣ�ֵΪ�������ݶ���
        /// </summary>
        private Dictionary<string, ImportDataInfo> dictImportData = new Dictionary<string, ImportDataInfo>();

        private ImportHandler()
        {
            this.loadData();
        }

        /// <summary>
        /// �������еĵ�����
        /// </summary>
        private void loadData()
        {
            //�������ļ����ж�ȡ�����ļ�����ϣ����
            string configXmlPath = AppDomain.CurrentDomain.BaseDirectory + Globals.ConfigXmlPath;
            string[] configFiles = System.IO.Directory.GetFiles(configXmlPath, "*.xml");

            foreach(string configFile in configFiles)
            {
                try
                {
                    ImportDataInfo importDataInfo = ConfigXmlLoader.GetDataInfoByXml(configFile);

                    //�����ϣ��
                    dictImportData[importDataInfo.TypeName] = importDataInfo;
                }
                catch (Exception ex)
                {
                    Console.Write("��ȡ" + configFile + "ʧ�ܣ�ԭ��" + ex.ToString());
                }
            }
        }

        /// <summary>
        /// �ؼ���
        /// </summary>
        public void reload()
        {
            dictImportData.Clear();
            this.loadData();
        }

        /// <summary>
        /// �õ����еĵ�������
        /// </summary>
        /// <returns></returns>
        public List<string> getAllImportType()
        {
            List<string> typeList = new List<string>();
            foreach (string type in dictImportData.Keys)
            {
                typeList.Add(type);
            }
            return typeList;
        }

        /// <summary>
        /// ���ݵ��������ַ�����Դ��������
        /// </summary>
        /// <param name="sourceTable"></param>
        /// <param name="importType"></param>
        public DataTable importData(DataTable sourceTable, string importType)
        {
            ImportDataInfo importDataInfo = this.getImportDataInfoByName(importType);
            if (importDataInfo != null)
            {
                return this.importData(sourceTable, importDataInfo);
            }
            else
            {
                throw new Exception("û�д����ͣ�");
            }
        }

        /// <summary>
        /// ͨ���������Ƶõ�������Ϣ(ImportDataInfo)
        /// </summary>
        /// <param name="importType"></param>
        /// <returns></returns>
        public ImportDataInfo getImportDataInfoByName(string importType)
        {
            ImportDataInfo importDataInfo = null;
            dictImportData.TryGetValue(importType, out importDataInfo);
            return importDataInfo;
        }


        /// <summary>
        /// ����DataTable������
        /// </summary>
        public DataTable importData(DataTable sourceTable, ImportDataInfo importDataInfo)
        {
            //��ɸѡ��Ҫ���뵽���ݿ����
            int primaryKeyCount = 0;        //���ݱ��а�������������
            List<AttributeInfo> attributeList = new List<AttributeInfo>();
            foreach (AttributeInfo attributeInfo in importDataInfo.AttributeList)
            {
                //�������Դ����������У��ͼ��뵼���е��б�
                if (sourceTable.Columns.Contains(attributeInfo.SourceName))
                {
                    attributeList.Add(attributeInfo);
                    //��������б����������������������һ
                    if (attributeInfo.IsPrimaryKey)
                    {
                        primaryKeyCount++;
                    }
                }
            }

            bool containsAllPrimaryKey = (primaryKeyCount == importDataInfo.PrimaryKeyList.Count);

            //��������Ŀ�����͵õ�DataManager,Ȼ����
            IDataManager dataManager = Fatory.GetDataManagerByType(importDataInfo.DestinationType);
            DataTable badRecords = dataManager.importData(sourceTable, attributeList, importDataInfo, containsAllPrimaryKey);
            return badRecords;
        }

        /// <summary>
        /// �õ�ʾ�����ݣ�����������ͣ���ȫ���ַ�����(��ʱֻʵ���˺�������)
        /// </summary>
        /// <param name="importType"></param>
        /// <param name="egnoreType"></param>
        /// <returns></returns>
        public DataTable getExampleData(string importType, bool egnoreType)
        {
            ImportDataInfo importDataInfo = this.getImportDataInfoByName(importType);
            if (importType == null)
            {
                return null;
            }

            //����ʾ�����ݱ�
            DataTable exampleTable = new DataTable(importDataInfo.TypeName);
            //����һ��
            exampleTable.Rows.Add(exampleTable.NewRow());

            foreach(AttributeInfo attributeInfo in importDataInfo.AttributeList)
            {
                if (!exampleTable.Columns.Contains(attributeInfo.SourceName))
                {
                    DataColumn column = new DataColumn(attributeInfo.SourceName);
                    if (egnoreType)
                    {
                        column.DataType = typeof(string);
                    }
                    exampleTable.Columns.Add(column);
                    exampleTable.Rows[0][column.ColumnName] = attributeInfo.ExampleData;
                }
            }
            return exampleTable;
        }
    }
}
