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

        /// <summary>
        /// 根据格式字符串填入attributeInfo（或determineColumnInfo）的columnName添加到sb中
        /// </summary>
        /// <param name="attributeInfo"></param>
        /// <param name="formatString">只能有{0}，就是ColumnName</param>
        /// <param name="sb"></param>
        private bool fillSqlByAttrWithDetermineColumn(AttributeInfo attributeInfo, string formatString, SqlStrType sqlStrType, StringBuilder sb)
        {
            bool hasPrimaryKey = false;
            this.fillSqlBySqlStrType(attributeInfo, formatString, sqlStrType, sb);
            if (attributeInfo.IsPrimaryKey)
            {
                hasPrimaryKey = true;
            }

            //遍历DetermineColumnList
            if (attributeInfo.DetermineColumnList != null)
            {
                foreach (DetermineColumnInfo determineColumnInfo in attributeInfo.DetermineColumnList)
                {
                    this.fillSqlBySqlStrType(determineColumnInfo, formatString, sqlStrType, sb);
                    if (determineColumnInfo.IsPrimaryKey)
                    {
                        hasPrimaryKey = true;
                    }
                }
            }
            return hasPrimaryKey;
        }

        /// <summary>
        /// 根据Sql片段的类型来向sb中填充
        /// </summary>
        /// <param name="basAttributeInfo"></param>
        /// <param name="formatString"></param>
        /// <param name="sqlStrType"></param>
        /// <param name="sb"></param>
        private void fillSqlBySqlStrType(BaseAttributeInfo basAttributeInfo, string formatString, SqlStrType sqlStrType, StringBuilder sb)
        {
            switch (sqlStrType)
            {
                case SqlStrType.ColumnName:
                    sb.Append(string.Format(formatString, basAttributeInfo.ColumnName));
                    break;
                case SqlStrType.ParamName:
                    sb.Append(string.Format(formatString, GetAttributeParameterSql(basAttributeInfo)));
                    break;
                case SqlStrType.ColumnAndParamName:
                    sb.Append(string.Format(formatString, basAttributeInfo.ColumnName, GetAttributeParameterSql(basAttributeInfo)));
                    break;
                default:
                    break;
            }
        }


        // <summary>
        /// 根据格式字符串填入attributeInfo（或determineColumnInfo）的columnName添加到sb中
        /// </summary>
        /// <param name="attributeInfo"></param>
        /// <param name="formatString">只能有{0}，就是ColumnName</param>
        /// <param name="sb"></param>
        private void fillSqlByAttrWithDetermineColumn(AttributeInfo attributeInfo, string formatString, SqlStrType sqlStrType, StringBuilder sb, ref bool hasPrimaryKey)
        {
            if (this.fillSqlByAttrWithDetermineColumn(attributeInfo, formatString, sqlStrType, sb))
            {
                hasPrimaryKey = true;
            }
        }

        
        #endregion

        /// <summary>
        /// 传入此方法的参数
        /// </summary>
        enum SqlStrType
        {
            /// <summary>
            /// 只需要传入ColumnName
            /// </summary>
            ColumnName,

            /// <summary>
            /// 只需要传入ParamName
            /// </summary>
            ParamName,

            /// <summary>
            /// ColumnName和ParamName都需要传入
            /// </summary>
            ColumnAndParamName
        }

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
        /// 使用insert导入数据
        /// </summary>
        /// <param name="sourceTable"></param>
        /// <param name="attributeList"></param>
        /// <param name="importDataInfo"></param>
        private DataTable importDataByInsert(DataTable sourceTable, List<AttributeInfo> attributeList, ImportDataInfo importDataInfo)
        {
            /**Insert语句
             * 示例：insert into table (column1, column2,....)
             * select :column1, :column2,.... from dual
             * where not exists (select 1 from tableName where primaryKey1 = :primaryKey1 and .....) 
             */

            StringBuilder sbInsert = new StringBuilder("insert into " + importDataInfo.TableName + " ( ");
            //第一个循环顺便寻找有没有主键
            bool hasPrimaryKey = false;
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                sbInsert.Append(attributeInfo.ColumnName + ", ");
                if (attributeInfo.IsPrimaryKey)
                {
                    hasPrimaryKey = true;
                }
            }
            sbInsert.Remove(sbInsert.Length - 2, 2);

            sbInsert.Append(" ) select ");
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                sbInsert.Append(GetAttributeParameterSql(attributeInfo) + ", ");
            }
            sbInsert.Remove(sbInsert.Length - 2, 2);
            sbInsert.Append(" from dual ");

            if (hasPrimaryKey)
            {
                //有主键，添加主键判断条件
                sbInsert.Append(string.Format(" where not exists (select 1 from {0} where ", importDataInfo.TableName));

                foreach (AttributeInfo attributeInfo in attributeList)
                {
                    if (attributeInfo.IsPrimaryKey)
                    {
                        fillSqlByAttrWithDetermineColumn(attributeInfo, " {0} = {1} and ", SqlStrType.ColumnAndParamName, sbInsert);
                    }
                }
                sbInsert.Remove(sbInsert.Length - 4, 4);
                sbInsert.Append(" ) ");
            }

            SqlConnection conn = new SqlConnection(GetConnectionString(importDataInfo));
            SqlCommand comm = new SqlCommand(sbInsert.ToString(), conn);

            //用一个与源表一样的表来存放没有更新成功的数据
            DataTable dtBadRecords = CreateTableByAttrList(sourceTable, attributeList);

            try
            {
                conn.Open();
                SqlParameter[] paras = new SqlParameter[attributeList.Count];
                for (int i = 0; i < attributeList.Count; i++)
                {
                    AttributeInfo attributeInfo = attributeList[i];
                    paras[i] = new SqlParameter(ParameterChar+attributeInfo.ColumnName, GetAttributeDbType(attributeInfo));//////by haoqiang
                    comm.Parameters.Add(paras[i]);
                }

                //逐行插入
                foreach (DataRow dataRow in sourceTable.Rows)
                {
                    try
                    {
                        //给parameter赋值
                        for (int i = 0; i < attributeList.Count; i++)
                        {
                            AttributeInfo attributeInfo = attributeList[i];
                            paras[i].Value = GetAttributeParameterValue(dataRow, attributeInfo);
                        }

                        comm.ExecuteNonQuery();
                    }
                    catch (SqlException ex)
                    {
                        AddNewRow(dtBadRecords, dataRow);
                    }
                }
            }
            catch (SqlException ex)
            {
                //说明open就出错了
                throw ex;
            }
            finally
            {
                conn.Close();
            }
            return dtBadRecords;
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
        /// 得到该列的Parameter表示sql,一般是:columnName
        /// </summary>
        /// <param name="attributeInfo"></param>
        /// <returns></returns>
        private static string GetAttributeParameterSql(BaseAttributeInfo baseAttributeInfo)
        {
            if (!string.IsNullOrEmpty(baseAttributeInfo.OuterTable))
            {
                string paramSql = string.Format("(select {0} from {1} where {2} = {3}{4} )", baseAttributeInfo.ValueColumn, baseAttributeInfo.OuterTable,
                                                baseAttributeInfo.KeyColumn, ParameterChar, baseAttributeInfo.ColumnName);
                return paramSql;
            }
            return ParameterChar + baseAttributeInfo.ColumnName;
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

        /// <summary>
        /// 得到该列的Parameter的值
        /// </summary>
        /// <param name="dataRow"></param>
        /// <param name="attributeInfo"></param>
        /// <returns></returns>
        private static object GetAttributeParameterValue(DataRow dataRow, DetermineColumnInfo determineColumnInfo)
        {
            try
            {
                string strData = determineColumnInfo.DetermineValue;

                if (string.IsNullOrEmpty(strData))
                {
                    //在oracle里！空字符串就是null！！！
                    return DBNull.Value;
                }

                //将日期型的转成DateTime
                if (determineColumnInfo.FormatStrings != null && determineColumnInfo.FormatStrings.Length > 0 && determineColumnInfo.ValueType.ToLower().Equals("date"))
                {
                    if (DataHelper.IsInvalidDate(strData))
                    {
                        return DBNull.Value;
                    }
                    foreach (string formatString in determineColumnInfo.FormatStrings)
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
                switch (determineColumnInfo.ValueType.ToLower())
                {
                    case "decimal":
                    case "int32":
                    case "int":
                    case "double":
                        return (strData.IndexOf('.') >= 0) ? (object)double.Parse(strData) : (object)int.Parse(strData);
                    case "date":
                        return DateTime.Parse(strData);
                    case "varchar":
                        if (determineColumnInfo.NeedEncrypt)
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
                throw new Exception(determineColumnInfo.ColumnName + "出错：" + ex.Message, ex);
            }
        }

        /// <summary>
        /// 得到连接字符串
        /// </summary>
        /// <param name="importDataInfo"></param>
        /// <returns></returns>
        private static string GetConnectionString(ImportDataInfo importDataInfo)
        {
            return string.IsNullOrEmpty(importDataInfo.ConnectionString) ? Globals.DefaultConnectionString : importDataInfo.ConnectionString;
        }
    }
}
