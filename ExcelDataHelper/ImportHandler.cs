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
        #region Singleton模式
        private static ImportHandler instance = new ImportHandler();

        public static ImportHandler GetInstance()
        {
            return instance;
        }
        #endregion

        /// <summary>
        /// 导入数据的哈希表，键为中文名称，值为导入数据对象
        /// </summary>
        private Dictionary<string, ImportDataInfo> dictImportData = new Dictionary<string, ImportDataInfo>();

        private ImportHandler()
        {
            this.loadData();
        }

        /// <summary>
        /// 加载所有的导入项
        /// </summary>
        private void loadData()
        {
            //从配置文件夹中读取配置文件到哈希表中
            string configXmlPath = AppDomain.CurrentDomain.BaseDirectory + Globals.ConfigXmlPath;
            string[] configFiles = System.IO.Directory.GetFiles(configXmlPath, "*.xml");

            foreach(string configFile in configFiles)
            {
                try
                {
                    ImportDataInfo importDataInfo = ConfigXmlLoader.GetDataInfoByXml(configFile);

                    //加入哈希表
                    dictImportData[importDataInfo.TypeName] = importDataInfo;
                }
                catch (Exception ex)
                {
                    Console.Write("读取" + configFile + "失败！原因：" + ex.ToString());
                }
            }
        }

        /// <summary>
        /// 重加载
        /// </summary>
        public void reload()
        {
            dictImportData.Clear();
            this.loadData();
        }

        /// <summary>
        /// 得到所有的导入类型
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
        /// 根据导入类型字符串和源表导入数据
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
                throw new Exception("没有此类型！");
            }
        }

        /// <summary>
        /// 通过类型名称得到导入信息(ImportDataInfo)
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
        /// 导入DataTable的数据
        /// </summary>
        public DataTable importData(DataTable sourceTable, ImportDataInfo importDataInfo)
        {
            //先筛选出要导入到数据库的列
            int primaryKeyCount = 0;        //数据表中包含的主键个数
            List<AttributeInfo> attributeList = new List<AttributeInfo>();
            foreach (AttributeInfo attributeInfo in importDataInfo.AttributeList)
            {
                //如果数据源表里包括该列，就加入导入列的列表
                if (sourceTable.Columns.Contains(attributeInfo.SourceName))
                {
                    attributeList.Add(attributeInfo);
                    //如果属性列表里包括主键列名，个数加一
                    if (attributeInfo.IsPrimaryKey)
                    {
                        primaryKeyCount++;
                    }
                }
            }

            bool containsAllPrimaryKey = (primaryKeyCount == importDataInfo.PrimaryKeyList.Count);

            //根据数据目标类型得到DataManager,然后导入
            IDataManager dataManager = Fatory.GetDataManagerByType(importDataInfo.DestinationType);
            DataTable badRecords = dataManager.importData(sourceTable, attributeList, importDataInfo, containsAllPrimaryKey);
            return badRecords;
        }

        /// <summary>
        /// 得到示例数据，如果忽略类型，则全是字符串型(暂时只实现了忽略类型)
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

            //创建示例数据表
            DataTable exampleTable = new DataTable(importDataInfo.TypeName);
            //加入一行
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
