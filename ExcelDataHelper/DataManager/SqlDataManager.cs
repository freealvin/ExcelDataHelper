using System;
using System.Collections.Generic;
using System.Text;
using ExcelDataHelper.DataManager;
using ExcelDataHelper.Entity;
using System.Data;
using System.Data.SqlClient;
using ExcelDataHelper.Helper;

namespace ExcelDataHelper.DataManager
{
    class SqlDataManager : IDataManager
    {
        private static string ParameterChar = "@";

        /// <summary>
        /// 得到默认的列类型
        /// </summary>
        /// <returns></returns>
        public string getDefaultColumnType()
        {
            return "varchar";
        }

        #region 使用导入多行的方式导入数据库
        /// <summary>
        /// 加上OracleParameter并且根据dataRow赋值，并且判断是否有主键但是没有值的
        /// </summary>
        /// <param name="dataRow"></param>
        /// <param name="attributeInfo"></param>
        /// <param name="paramList"></param>
        private bool addParameterWithValue(DataRow dataRow, AttributeInfo attributeInfo, System.Collections.IList paramList)
        {
            bool hasNullPrimary = false;
            SqlParameter tempParm = new SqlParameter(attributeInfo.ColumnName, GetAttributeDbType(attributeInfo));
            tempParm.Value = GetAttributeParameterValue(dataRow, attributeInfo);
            if (attributeInfo.IsPrimaryKey && tempParm.Value.Equals(DBNull.Value))
            {
                hasNullPrimary = true;
            }
            paramList.Add(tempParm);

            //遍历被确定列的列表
            if (attributeInfo.DetermineColumnList != null)
            {
                foreach (DetermineColumnInfo determineColumnInfo in attributeInfo.DetermineColumnList)
                {
                    tempParm = new SqlParameter(determineColumnInfo.ColumnName, GetAttributeDbType(determineColumnInfo));
                    tempParm.Value = GetAttributeParameterValue(dataRow, determineColumnInfo);
                    if (determineColumnInfo.IsPrimaryKey && tempParm.Value.Equals(DBNull.Value))
                    {
                        hasNullPrimary = true;
                    }
                    paramList.Add(tempParm);
                }
            }
            return hasNullPrimary;
        }
        #endregion

        /// <summary>
        /// 根据源表和attributeList创建DataTable，并添加所有attributeList和源表里都包含的列
        /// </summary>
        /// <param name="sourceTable"></param>
        /// <param name="attributeList"></param>
        /// <returns></returns>
        private static DataTable CreateTableByAttrList(DataTable sourceTable, List<AttributeInfo> attributeList)
        {
            DataTable newTable = new DataTable();
            //添加所有attributeList里包含的列
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                if (sourceTable.Columns.Contains(attributeInfo.SourceName) && !newTable.Columns.Contains(attributeInfo.SourceName))
                {
                    DataColumn sourceColumn = sourceTable.Columns[attributeInfo.SourceName];
                    newTable.Columns.Add(new DataColumn(sourceColumn.ColumnName, sourceColumn.DataType, sourceColumn.Expression, sourceColumn.ColumnMapping));
                }
            }
            newTable.Clear();
            return newTable;
        }

        /// <summary>
        /// 往一个表里加入一行
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="dr"></param>
        private static DataRow AddNewRow(DataTable destTable, DataRow sourceRow)
        {
            DataRow newRow = destTable.NewRow();

            foreach (DataColumn column in destTable.Columns)
            {
                newRow[column.ColumnName] = sourceRow[column.ColumnName];
            }
            destTable.Rows.Add(newRow);
            return destTable.Rows[destTable.Rows.Count - 1];
        }
        
        /// <summary>
        /// 得到该列Parameter的类型
        /// </summary>
        /// <param name="attributeInfo"></param>
        /// <returns></returns>
        private static SqlDbType GetAttributeDbType(BaseAttributeInfo baseAttributeInfo)
        {
            SqlDbType dbType = (SqlDbType)Enum.Parse(typeof(SqlDbType), baseAttributeInfo.ValueType, true);
            return dbType;
        }

        /// <summary>
        /// 得到该列的Parameter的值,如果为null或空字符串，是字符串型则
        /// </summary>
        /// <param name="dataRow"></param>
        /// <param name="attributeInfo"></param>
        /// <returns></returns>
        private static object GetAttributeParameterValue(DataRow dataRow, AttributeInfo attributeInfo)
        {
            try
            {
                object objData = dataRow[attributeInfo.SourceName];
                string strData = (objData == null || objData.Equals(DBNull.Value)) ? null : objData.ToString();
                if (string.IsNullOrEmpty(strData))
                {
                    return DBNull.Value;
                }

                //将前后空格去掉
                if (attributeInfo.Trim)
                {
                    strData = strData.Trim();
                    if (string.IsNullOrEmpty(strData))
                    {
                        ////去掉空格后为空字符串，如果是字符串型，则返回字符串，如果不是，则返回空
                        //switch (attributeInfo.AttributeType.ToLower())
                        //{
                        //    case "varchar2":
                        //        return strData;
                        //    default:
                        //        return DBNull.Value;
                        //}

                        //在oracle里！空字符串就是null！！！
                        return DBNull.Value;
                    }
                }
                //取出确定值
                if (attributeInfo.IsCertainValue)
                {
                    bool isCertainValue = false;
                    foreach (CertainValueInfo certainValue in attributeInfo.CertainValueList)
                    {
                        if (strData.Equals(certainValue.Key))
                        {
                            isCertainValue = true;
                            strData = certainValue.Value;
                            break;
                        }
                    }
                    if (!isCertainValue)
                    {
                        throw new Exception(attributeInfo.SourceName + "的值不在规定的值中！");
                    }
                }
                //将日期型的转成DateTime
                if (attributeInfo.FormatStrings != null && attributeInfo.FormatStrings.Length > 0 && attributeInfo.ValueType.ToLower().Equals("date"))
                {
                    if (strData.Equals("000000"))
                    {
                        return DBNull.Value;
                    }
                    foreach (string formatString in attributeInfo.FormatStrings)
                    {
                        try
                        {
                            return DateTime.ParseExact(strData, formatString, null);
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                }

                //根据类型返回
                switch (attributeInfo.ValueType.ToLower())
                {
                    case "decimal":
                    case "int32":
                    case "double":
                        return (strData.IndexOf('.') >= 0) ? (object)double.Parse(strData) : (object)int.Parse(strData);
                    case "date":
                        return DateTime.Parse(strData);
                    case "varchar":
                        if (attributeInfo.NeedEncrypt)
                        {
                            strData = EncryptHelper.EncrptString(strData);
                        }
                        return strData;
                    default:
                        return strData;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(attributeInfo.SourceName + "出错：" + ex.Message, ex);
            }
        }
    }
}
