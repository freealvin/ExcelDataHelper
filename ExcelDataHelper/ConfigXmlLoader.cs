using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using DataImportLib.Entity;
using DataImportLib.DataManager;

namespace DataImportLib
{
    class ConfigXmlLoader
    {
        /// <summary>
        /// �������ļ���õ�һ���������ݵ���Ϣ��
        /// </summary>
        /// <param name="configFile"></param>
        /// <returns></returns>
        public static ImportDataInfo GetDataInfoByXml(string configFile)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(configFile);

            //��ʱ����ַ���������
            string tempAttrStr = null;

            //��ȡdata�������Ը�ֵ
            XmlNode dataNode = xmlDoc.SelectSingleNode("data");
            ImportDataInfo importDataInfo = new ImportDataInfo();
            importDataInfo.TypeName = dataNode.Attributes["typeName"].Value;
            importDataInfo.TableName = dataNode.Attributes["tableName"].Value;
            importDataInfo.DestinationType = (DestinationType)Enum.Parse(typeof(DestinationType), dataNode.Attributes["destinationType"].Value, true);

            if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(dataNode.Attributes, "connectionString", true)))
            {
                importDataInfo.ConnectionString = tempAttrStr;
            }

            if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(dataNode.Attributes, "autoMerge", true)))
            {
                importDataInfo.AutoMerge = bool.Parse(tempAttrStr);
            }
            
            if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(dataNode.Attributes, "onlyUpdate", true)))
            {
                importDataInfo.OnlyUpdate = bool.Parse(tempAttrStr);
            }

            if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(dataNode.Attributes, "oneToMultiRow", true)))
            {
                importDataInfo.OneToMultiRow = bool.Parse(tempAttrStr);
            }

            //��������Ŀ�����͵õ�DataManager
            IDataManager dataManager = Fatory.GetDataManagerByType(importDataInfo.DestinationType);

            //��֤���е����Ƿ�ֻ��һ��
            Dictionary<string, object> dictColumn = new Dictionary<string, object>();
            //��ȡattr����б�
            List<AttributeInfo> attributeList = new List<AttributeInfo>();
            List<AttributeInfo> primaryKeyList = new List<AttributeInfo>();
            XmlNodeList attrNodeList = dataNode.SelectNodes("attr");
            foreach (XmlNode attrNode in attrNodeList)
            {
                AttributeInfo attributeInfo = new AttributeInfo();
                attributeInfo.ColumnName = attrNode.Attributes["columnName"].Value;
                if (importDataInfo.OneToMultiRow)
                {
                    //��֤���е����Ƿ�ֻ��һ��
                    if (!dictColumn.ContainsKey(attributeInfo.ColumnName))
                    {
                        dictColumn.Add(attributeInfo.ColumnName, null);
                    }
                    else
                    {
                        //�Ѿ����ֹ��������
                        if (string.IsNullOrEmpty(importDataInfo.MultiRowColumnName))
                        {
                            importDataInfo.MultiRowColumnName = attributeInfo.ColumnName;
                        }
                        else if (!importDataInfo.MultiRowColumnName.Equals(attributeInfo.ColumnName))
                        {
                            throw new Exception("�ж���г����˶�Σ�(ֻ����һ���г��ֶ��)");
                        }
                    }
                }

                attributeInfo.SourceName = attrNode.Attributes["sourceName"].Value;

                if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(attrNode.Attributes, "isCertainValue", true)))
                {
                    attributeInfo.IsCertainValue = bool.Parse(tempAttrStr);

                    //�����ȷ����ֵ��Χ�������ȷ��ֵ�б�
                    if (attributeInfo.IsCertainValue)
                    {
                        XmlNode certainValuesNode = attrNode.SelectSingleNode("certainValues");
                        XmlNodeList certainValueNodeList = certainValuesNode.SelectNodes("certainValue");
                        List<CertainValueInfo> certainValueList = new List<CertainValueInfo>();

                        if (certainValueNodeList != null)
                        {
                            foreach (XmlNode certainValueNode in certainValueNodeList)
                            {
                                CertainValueInfo certainValueInfo = new CertainValueInfo();
                                certainValueInfo.Key = certainValueNode.Attributes["key"].Value;
                                certainValueInfo.Value = certainValueNode.Attributes["value"].Value;

                                certainValueList.Add(certainValueInfo);
                            }
                        }
                        attributeInfo.CertainValueList = certainValueList;
                    }
                }

                //�����ȷ�����У������ȷ�����б�
                XmlNode determineColumnsNode = attrNode.SelectSingleNode("determineColumns");
                if (determineColumnsNode != null)
                {
                    XmlNodeList determineColumnNodeList = determineColumnsNode.SelectNodes("determineColumn");
                    if (determineColumnNodeList != null)
                    {
                        List<DetermineColumnInfo> determineColumnList = new List<DetermineColumnInfo>();
                        foreach (XmlNode determineColumnNode in determineColumnNodeList)
                        {
                            DetermineColumnInfo determineColumnInfo = new DetermineColumnInfo();
                            /**determineColumnInfo.ColumnName = determineColumnNode.Attributes["columnName"].Value;

                            //�������͵ĸ�ֵ�����Ϊ������dataManager��Ĭ��ֵ
                            if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(determineColumnNode.Attributes, "valueType", true)))
                            {
                                determineColumnInfo.ValueType = tempAttrStr;
                            }
                            else
                            {
                                determineColumnInfo.ValueType = dataManager.getDefaultColumnType();
                            }

                            if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(determineColumnNode.Attributes, "determineValue", true)))
                            {
                                determineColumnInfo.DetermineValue = tempAttrStr;
                            }

                            if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(determineColumnNode.Attributes, "isPrimaryKey", true)))
                            {
                                determineColumnInfo.IsPrimaryKey = bool.Parse(tempAttrStr);
                            }

                            if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(determineColumnNode.Attributes, "outerTable", true)))
                            {
                                determineColumnInfo.OuterTable = tempAttrStr;
                            }

                            if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(determineColumnNode.Attributes, "keyColumn", true)))
                            {
                                determineColumnInfo.KeyColumn = tempAttrStr;
                            }

                            if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(determineColumnNode.Attributes, "valueColumn", true)))
                            {
                                determineColumnInfo.ValueColumn = tempAttrStr;
                            }

                            if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(determineColumnNode.Attributes, "formatString", true)))
                            {
                                determineColumnInfo.FormatStrings = tempAttrStr.Split(Globals.FormatStringSplitter);
                            }*/
                            FillBaseAttributeInfo(determineColumnInfo, determineColumnNode, dataManager);

                            determineColumnList.Add(determineColumnInfo);
                        }
                        attributeInfo.DetermineColumnList = determineColumnList;
                    }
                }

                if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(attrNode.Attributes, "trim", true)))
                {
                    attributeInfo.Trim = bool.Parse(tempAttrStr);
                }

                /*//�������͵ĸ�ֵ�����Ϊ������dataManager��Ĭ��ֵ
                if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(attrNode.Attributes, "valueType", true)))
                {
                    attributeInfo.ValueType = tempAttrStr;
                }
                else
                {
                    attributeInfo.ValueType = dataManager.getDefaultColumnType();
                }

                if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(attrNode.Attributes, "isPrimaryKey", true)))
                {
                    attributeInfo.IsPrimaryKey = bool.Parse(tempAttrStr);
                }

                if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(attrNode.Attributes, "outerTable", true)))
                {
                    attributeInfo.OuterTable = tempAttrStr;
                }

                if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(attrNode.Attributes, "keyColumn", true)))
                {
                    attributeInfo.KeyColumn = tempAttrStr;
                }

                if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(attrNode.Attributes, "valueColumn", true)))
                {
                    attributeInfo.ValueColumn = tempAttrStr;
                }

                if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(attrNode.Attributes, "formatString", true)))
                {
                    attributeInfo.FormatStrings = tempAttrStr.Split(Globals.FormatStringSplitter);
                }*/
                FillBaseAttributeInfo(attributeInfo, attrNode, dataManager);

                if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(attrNode.Attributes, "exampleData", true)))
                {
                    attributeInfo.ExampleData = tempAttrStr;
                }
                if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(attrNode.Attributes, "needEncrypt", true)))
                {
                    attributeInfo.NeedEncrypt = bool.Parse(tempAttrStr);
                }
                
                attributeList.Add(attributeInfo);
                if (attributeInfo.IsPrimaryKey)
                {
                    primaryKeyList.Add(attributeInfo);
                }
            }
            importDataInfo.AttributeList = attributeList;
            importDataInfo.PrimaryKeyList = primaryKeyList;

            return importDataInfo;
        }

        /// <summary>
        /// ��XmlNode�����ȡ��Ϣ���BaseAttributeInfo
        /// </summary>
        /// <param name="baseAttributeInfo"></param>
        /// <param name="baseAttributeNode"></param>
        /// <param name="dataManager"></param>
        public static void FillBaseAttributeInfo(BaseAttributeInfo baseAttributeInfo, XmlNode baseAttributeNode, IDataManager dataManager)
        {
            string tempAttrStr = null;
            baseAttributeInfo.ColumnName = baseAttributeNode.Attributes["columnName"].Value;

            //�������͵ĸ�ֵ�����Ϊ������dataManager��Ĭ��ֵ
            if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(baseAttributeNode.Attributes, "valueType", true)))
            {
                baseAttributeInfo.ValueType = tempAttrStr;
            }
            else
            {
                baseAttributeInfo.ValueType = dataManager.getDefaultColumnType();
            }

            if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(baseAttributeNode.Attributes, "determineValue", true)))
            {
                baseAttributeInfo.DetermineValue = tempAttrStr;
            }

            if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(baseAttributeNode.Attributes, "isPrimaryKey", true)))
            {
                baseAttributeInfo.IsPrimaryKey = bool.Parse(tempAttrStr);
            }

            if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(baseAttributeNode.Attributes, "outerTable", true)))
            {
                baseAttributeInfo.OuterTable = tempAttrStr;
            }

            if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(baseAttributeNode.Attributes, "keyColumn", true)))
            {
                baseAttributeInfo.KeyColumn = tempAttrStr;
            }

            if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(baseAttributeNode.Attributes, "valueColumn", true)))
            {
                baseAttributeInfo.ValueColumn = tempAttrStr;
            }

            if (!string.IsNullOrEmpty(tempAttrStr = GetStringAttribute(baseAttributeNode.Attributes, "formatString", true)))
            {
                baseAttributeInfo.FormatStrings = tempAttrStr.Split(Globals.FormatStringSplitter);
            }
        }

        /// <summary>
        /// �õ����Լ�������ַ���ֵ������ѡ����ַ����Ƿ���ʾΪ��
        /// </summary>
        /// <param name="attributes"></param>
        /// <param name="attrName"></param>
        /// <param name="isEmptyNull">���Ϊ���ַ����Ƿ񷵻�null</param>
        /// <returns></returns>
        public static string GetStringAttribute(XmlAttributeCollection attributes, string attrName, bool isEmptyNull)
        {
            XmlAttribute tempXmlAttr = null;
            if ((tempXmlAttr = attributes[attrName]) != null)
            {
                if (isEmptyNull && string.IsNullOrEmpty(tempXmlAttr.Value))
                {
                    return null;
                }
                return tempXmlAttr.Value;
            }
            return null;
        }
    }
}
