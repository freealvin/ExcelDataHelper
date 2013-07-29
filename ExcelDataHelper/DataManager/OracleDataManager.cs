using System;
using System.Collections.Generic;
using System.Text;
using ExcelDataHelper.Entity;
using System.Data;
using Oracle.DataAccess.Client;
using ExcelDataHelper.Helper;

namespace ExcelDataHelper.DataManager
{
    /// <summary>
    /// Oracle的数据管理器
    /// </summary>
    class OracleDataManager : IDataManager
    {
        private static string ParameterChar = ":";
        /// <summary>
        /// 得到默认的列类型
        /// </summary>
        /// <returns></returns>
        public string getDefaultColumnType()
        {
            return "varchar2";
        }

        /// <summary>
        /// 根据数据源表和要导入的列信息导入目的
        /// </summary>
        /// <param name="sourceTable"></param>
        /// <param name="attributeList"></param>
        /// <param name="importDataInfo"></param>
        public DataTable importData(DataTable sourceTable, List<AttributeInfo> attributeList, ImportDataInfo importDataInfo, bool containsAllPrimaryKey)
        {
            if (importDataInfo.OneToMultiRow)
            {
                return this.imoprtDataToMultiRow(sourceTable, attributeList, importDataInfo);
            }
            else if (importDataInfo.OnlyUpdate)
            {
                return this.importDataByUpdate(sourceTable, attributeList, importDataInfo);
            }
            else if (importDataInfo.AutoMerge && containsAllPrimaryKey && attributeList.Count > importDataInfo.PrimaryKeyList.Count)
            {
                return this.importDataByMerge(sourceTable, attributeList, importDataInfo);
            }
            else
            {
                return this.importDataByInsert(sourceTable, attributeList, importDataInfo);
            }
        }

        #region 使用导入多行的方式导入数据库
        /// <summary>
        /// 使用导入多行的方式导入数据库
        /// </summary>
        /// <param name="sourceTable"></param>
        /// <param name="attributeList"></param>
        /// <param name="importDataInfo"></param>
        /// <returns></returns>
        private DataTable imoprtDataToMultiRow(DataTable sourceTable, List<AttributeInfo> attributeList, ImportDataInfo importDataInfo)
        {
            /**
             * 带主键检查的Insert语句
             * insert into tableName(column1, column2, column3......multiColumn, determineColumn1...)
             * select :column1, :column2, :column3......:multiColumn from dual
             * where not exists (select 1 from tableName where primaryKey1 = :primaryKey1 and .....) 
             */

            bool hasPrimaryKey = false;

            #region 构造语句列表

            StringBuilder sbInsert = new StringBuilder();
            sbInsert.Append(string.Format("insert into {0} (", importDataInfo.TableName));

            //先遍历不是多行列的列
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                if (!attributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName))
                {
                    this.fillSqlByAttrWithDetermineColumn(attributeInfo, "{0}, ", SqlStrType.ColumnName, sbInsert, ref hasPrimaryKey);
                }
            }

            //再遍历是多行列的，创建insert语句列表
            List<StringBuilder> sbInsertList = new List<StringBuilder>();
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                if (attributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName))
                {
                    StringBuilder sbNewInsert = new StringBuilder();
                    sbNewInsert.Append(sbInsert.ToString());
                    this.fillSqlByAttrWithDetermineColumn(attributeInfo, "{0}, ", SqlStrType.ColumnName, sbNewInsert, ref hasPrimaryKey);

                    sbInsertList.Add(sbNewInsert);
                }
            }

            //移除最后的", "，再继续填下一个
            foreach (StringBuilder sbNewInsert in sbInsertList)
            {
                sbNewInsert.Remove(sbNewInsert.Length - 2, 2);
                sbNewInsert.Append(" ) select ");
            }

            //select字段列表
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                if (!attributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName))
                {
                    foreach (StringBuilder sbNewInsert in sbInsertList)
                    {
                        this.fillSqlByAttrWithDetermineColumn(attributeInfo, "{0}, ", SqlStrType.ParamName, sbNewInsert, ref hasPrimaryKey);
                    }
                }
            }

            //多行的列按顺序逐个添加
            int sbInsertIndex = 0;
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                if (attributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName))
                {
                    StringBuilder sbNewInsert = sbInsertList[sbInsertIndex++];
                    this.fillSqlByAttrWithDetermineColumn(attributeInfo, "{0}, ", SqlStrType.ParamName, sbNewInsert, ref hasPrimaryKey);
                }
            }

            //移除最后的", "，再继续填下一个
            foreach (StringBuilder sbNewInsert in sbInsertList)
            {
                sbNewInsert.Remove(sbNewInsert.Length - 2, 2);
                sbNewInsert.Append(" from dual ");
            }

            if (hasPrimaryKey)
            {
                foreach (StringBuilder sbNewInsert in sbInsertList)
                {
                    sbNewInsert.Append(string.Format(" where not exists (select 1 from {0} where ", importDataInfo.TableName));
                }

                //判断主键的where条件
                foreach (AttributeInfo attributeInfo in attributeList)
                {
                    if (attributeInfo.IsPrimaryKey && !attributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName))
                    {
                        foreach (StringBuilder sbNewInsert in sbInsertList)
                        {
                            this.fillSqlByAttrWithDetermineColumn(attributeInfo, " {0} = {1} and ", SqlStrType.ColumnAndParamName, sbNewInsert);
                        }
                    }
                }

                sbInsertIndex = 0;
                foreach (AttributeInfo attributeInfo in attributeList)
                {
                    if (attributeInfo.IsPrimaryKey && attributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName))
                    {
                        StringBuilder sbNewInsert = sbInsertList[sbInsertIndex++];
                        this.fillSqlByAttrWithDetermineColumn(attributeInfo, " {0} = {1} and ", SqlStrType.ColumnAndParamName, sbNewInsert);
                    }
                }
                //移除最后的"and "，再加上最后的右括号
                foreach (StringBuilder sbNewInsert in sbInsertList)
                {
                    sbNewInsert.Remove(sbNewInsert.Length - 4, 4);
                    sbNewInsert.Append(" )");
                }
            }
            #endregion

            OracleConnection conn = new OracleConnection(GetConnectionString(importDataInfo));
            OracleCommand comm = new OracleCommand();
            comm.Connection = conn;

            //用一个与源表一样的表来存放没有更新成功的数据
            DataTable dtBadRecords = CreateTableByAttrList(sourceTable, attributeList);

            try
            {
                conn.Open();

                #region 执行语句
                foreach (DataRow dataRow in sourceTable.Rows)
                {
                    //那些单行的列的参数列表
                    List<OracleParameter> singleRowParamList = new List<OracleParameter>();
                    //判断主键的where条件
                    bool hasNullPrimary = false;
                    foreach (AttributeInfo attributeInfo in attributeList)
                    {
                        if (!attributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName))
                        {
                            hasNullPrimary = this.addOracleParameterWithValue(dataRow, attributeInfo, singleRowParamList);
                            if (hasNullPrimary)
                            {
                                //此行不能被导入，因为单行列中有主键为空的
                                AddNewRow(dtBadRecords, dataRow);
                                break;
                            }
                        }
                    }

                    if (hasNullPrimary)
                    {
                        //下一行
                        continue;
                    }

                    int commandTextIndex = 0;
                    //多行列的参数，一列执行一次
                    foreach (AttributeInfo attributeInfo in attributeList)
                    {
                        if (attributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName))
                        {
                            comm.CommandText = sbInsertList[commandTextIndex++].ToString();
                            hasNullPrimary = true;
                            //先加入
                            comm.Parameters.Clear();
                            foreach (OracleParameter param in singleRowParamList)
                            {
                                comm.Parameters.Add(param);
                            }

                            hasNullPrimary = this.addOracleParameterWithValue(dataRow, attributeInfo, comm.Parameters);

                            if (hasNullPrimary)
                            {
                                //下一行
                                continue;
                            }

                            try
                            {
                                comm.ExecuteNonQuery();
                            }
                            catch (OracleException ex)
                            {
                                DataRow newRow = AddNewRow(dtBadRecords, dataRow);

                                //将其他多行列置空
                                foreach (AttributeInfo otherAttributeInfo in attributeList)
                                {
                                    if (otherAttributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName) && otherAttributeInfo != attributeInfo)
                                    {
                                        newRow[otherAttributeInfo.SourceName] = null;
                                    }

                                }
                            }
                        }
                    }
                }
                #endregion 
            }
            catch (OracleException ex)
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
        /// 加上OracleParameter并且根据dataRow赋值，并且判断是否有主键但是没有值的
        /// </summary>
        /// <param name="dataRow"></param>
        /// <param name="attributeInfo"></param>
        /// <param name="paramList"></param>
        private bool addOracleParameterWithValue(DataRow dataRow, AttributeInfo attributeInfo, System.Collections.IList paramList)
        {
            bool hasNullPrimary = false;
            OracleParameter tempParm = new OracleParameter(attributeInfo.ColumnName, GetAttributeDbType(attributeInfo));
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
                    tempParm = new OracleParameter(determineColumnInfo.ColumnName, GetAttributeDbType(determineColumnInfo));
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

        /// <summary>
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
        /// 根据源表和attributeList创建DataTable，并添加所有attributeList和源表里都包含的列
        /// </summary>
        /// <param name="sourceTable"></param>
        /// <param name="attributeList"></param>
        /// <returns></returns>
        private static DataTable CreateTableByAttrList(DataTable sourceTable, List<AttributeInfo> attributeList)
        {
            DataTable newTable = new DataTable();
            //添加所有attributeList里包含的列
            foreach(AttributeInfo attributeInfo in attributeList)
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

            OracleConnection conn = new OracleConnection(GetConnectionString(importDataInfo));
            OracleCommand comm = new OracleCommand(sbInsert.ToString(), conn);

            //用一个与源表一样的表来存放没有更新成功的数据
            DataTable dtBadRecords = CreateTableByAttrList(sourceTable, attributeList);

            try
            {
                conn.Open();
                OracleParameter[] paras = new OracleParameter[attributeList.Count];
                for (int i=0; i < attributeList.Count; i++)
                {
                    AttributeInfo attributeInfo = attributeList[i];
                    paras[i] = new OracleParameter(attributeInfo.ColumnName, GetAttributeDbType(attributeInfo));
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
                            paras[i].Value = GetAttributeParameterValue(dataRow, attributeInfo) ;
                        }

                        comm.ExecuteNonQuery();
                    }
                    catch (OracleException ex)
                    {
                        AddNewRow(dtBadRecords, dataRow);
                    }
                }
            }
            catch (OracleException ex)
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
        /// 使用Update导入数据
        /// </summary>
        /// <param name="sourceTable"></param>
        /// <param name="attributeList"></param>
        /// <param name="importDataInfo"></param>
        private DataTable importDataByUpdate(DataTable sourceTable, List<AttributeInfo> attributeList, ImportDataInfo importDataInfo)
        {
            /**
             * Update语句(因为其他写法加Parameter会有莫名其妙的bug)
             * 示例： update table t set column1 = :column1, column2 = :column2....
             *  where t.primaryKey1 = :primaryKey1 and t.primaryKey2 = :primaryKey2...
             */
            List<AttributeInfo> primaryKeyList = importDataInfo.PrimaryKeyList;
            if (primaryKeyList.Count == 0)
            {
                throw new Exception("使用更新方法必须有主键列");
            }

            StringBuilder sbMerge = new StringBuilder(string.Format(" update {0} t set ", importDataInfo.TableName));
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                //只update不是主键的列
                if (!attributeInfo.IsPrimaryKey)
                {
                    //"columnName = :columnName, "
                    sbMerge.Append(string.Format("{0} = {1}, ", attributeInfo.ColumnName, GetAttributeParameterSql(attributeInfo)));
                }
            }
            sbMerge.Remove(sbMerge.Length - 2, 2);

            sbMerge.Append(" where ");

            //where条件
            foreach (AttributeInfo attributeInfo in primaryKeyList)
            {
                //"t2.columnName = t.columnName and "
                sbMerge.Append(string.Format("t.{0} = {1} and ", attributeInfo.ColumnName, GetAttributeParameterSql(attributeInfo)));
            }
            sbMerge.Remove(sbMerge.Length - 4, 4);


            OracleConnection conn = new OracleConnection(GetConnectionString(importDataInfo));
            OracleCommand comm = new OracleCommand(sbMerge.ToString(), conn);

            //用一个与源表一样的表来存放没有更新成功的数据
            DataTable dtBadRecords = CreateTableByAttrList(sourceTable, attributeList);
            try
            {
                conn.Open();
                foreach (AttributeInfo attributeInfo in attributeList)
                {
                    if (!attributeInfo.IsPrimaryKey)
                    {
                        comm.Parameters.Add(
                            new OracleParameter(attributeInfo.ColumnName, GetAttributeDbType(attributeInfo)));
                    }
                }
                foreach (AttributeInfo primaryAttributeInfo in primaryKeyList)
                {
                    comm.Parameters.Add(
                        new OracleParameter(primaryAttributeInfo.ColumnName, GetAttributeDbType(primaryAttributeInfo)));
                }
                //逐行插入
                foreach (DataRow dataRow in sourceTable.Rows)
                {
                    try
                    {
                        //给parameter赋值
                        int paraPos = 0;
                        foreach (AttributeInfo attributeInfo in attributeList)
                        {
                            if (!attributeInfo.IsPrimaryKey)
                            {
                                comm.Parameters[paraPos++].Value = GetAttributeParameterValue(dataRow, attributeInfo);
                            }
                        }
                        foreach (AttributeInfo primaryAttributeInfo in primaryKeyList)
                        {
                            comm.Parameters[paraPos++].Value = GetAttributeParameterValue(dataRow, primaryAttributeInfo);
                        }
                        int row = comm.ExecuteNonQuery();
                        //如果没有影响的行说明也没有成功
                        if (row <= 0)
                        {
                            AddNewRow(dtBadRecords, dataRow);
                        }
                    }
                    catch (Exception ex)
                    {
                        AddNewRow(dtBadRecords, dataRow);
                    }
                }
            }
            catch (OracleException ex)
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
        /// 使用merge导入数据
        /// </summary>
        /// <param name="sourceTable"></param>
        /// <param name="attributeList"></param>
        /// <param name="importDataInfo"></param>
        private DataTable importDataByMerge(DataTable sourceTable, List<AttributeInfo> attributeList, ImportDataInfo importDataInfo)
        {
            /**
             * Merge语句(因为其他写法加Parameter会有莫名其妙的bug)
             * 示例：merge into table t
             * using ( select '1' c from dual ) t2 on (t2.c = '1' and t.primaryKey1 = :primaryKey1 and t.primaryKey2 = :primaryKey2...)
             * when matched then update set column1 = :column1, column2 = :column2....
             * when not matched then insert (column1, column2, .....) values (:column1, :column2,.......)
             */

            StringBuilder sbMerge = new StringBuilder(string.Format(" merge into {0} t ", importDataInfo.TableName));
            sbMerge.Append(" using (select '1' c from dual) t2 on (t2.c = '1' and ");
            List<AttributeInfo> primaryKeyList = importDataInfo.PrimaryKeyList;
            //t的on条件
            foreach (AttributeInfo attributeInfo in primaryKeyList)
            {
                //"t2.columnName = t.columnName and "
                sbMerge.Append(string.Format("t.{0} = {1}{0} and ", attributeInfo.ColumnName, ParameterChar));
            }
            sbMerge.Remove(sbMerge.Length - 4, 4);

            //update的set字段列表
            sbMerge.Append(" ) when matched then update set ");
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                //只update不是主键的列
                if (!attributeInfo.IsPrimaryKey)
                {
                    //"columnName = :columnName, "
                    sbMerge.Append(string.Format("{0} = {1}, ", attributeInfo.ColumnName, GetAttributeParameterSql(attributeInfo)));
                }
            }
            sbMerge.Remove(sbMerge.Length - 2, 2);

            //insert字段
            sbMerge.Append(" when not matched then insert (");
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                sbMerge.Append(attributeInfo.ColumnName + ", ");
            }
            sbMerge.Remove(sbMerge.Length - 2, 2);

            //insert值字段
            sbMerge.Append(" ) values (");
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                sbMerge.Append(GetAttributeParameterSql(attributeInfo) + ", ");
            }
            sbMerge.Remove(sbMerge.Length - 2, 2);
            sbMerge.Append(" )");


            OracleConnection conn = new OracleConnection(GetConnectionString(importDataInfo));
            OracleCommand comm = new OracleCommand(sbMerge.ToString(), conn);

            //用一个与源表一样的表来存放没有更新成功的数据
            DataTable dtBadRecords = CreateTableByAttrList(sourceTable, attributeList);

            try
            {
                conn.Open();
                foreach(AttributeInfo primaryAttributeInfo in primaryKeyList)
                {
                    comm.Parameters.Add(
                        new OracleParameter(primaryAttributeInfo.ColumnName, GetAttributeDbType(primaryAttributeInfo)));
                }
                foreach (AttributeInfo attributeInfo in attributeList)
                {
                    if (!attributeInfo.IsPrimaryKey)
                    {
                        comm.Parameters.Add(
                            new OracleParameter(attributeInfo.ColumnName, GetAttributeDbType(attributeInfo)));
                    }
                }
                //逐行插入
                foreach (DataRow dataRow in sourceTable.Rows)
                {
                    try
                    {
                        //给parameter赋值
                        int paraPos = 0;
                        foreach (AttributeInfo primaryAttributeInfo in primaryKeyList)
                        {
                            comm.Parameters[paraPos++].Value = GetAttributeParameterValue(dataRow, primaryAttributeInfo);
                        }

                        foreach (AttributeInfo attributeInfo in attributeList)
                        {
                            if (!attributeInfo.IsPrimaryKey)
                            {
                                if (attributeInfo.SourceName.StartsWith("学士"))
                                {
                                    ;
                                }

                                comm.Parameters[paraPos++].Value = GetAttributeParameterValue(dataRow, attributeInfo);
                            }
                        }

                        comm.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        AddNewRow(dtBadRecords, dataRow);
                    }
                }
            }
            catch (OracleException ex)
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
                    case "int":
                    case "double":
                        return (strData.IndexOf('.') >= 0) ? (object)double.Parse(strData) : (object)int.Parse(strData);
                    case "date":
                        return DateTime.Parse(strData);
                    case "varchar2":
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
                    case "double":
                        return (strData.IndexOf('.') >= 0) ? (object)double.Parse(strData) : (object)int.Parse(strData);
                    case "date":
                        return DateTime.Parse(strData);
                    case "varchar2":
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

#region 废代码
/****先生成没有多列的语句
            bool hasPrimaryKey = false;
            StringBuilder sbCountSql = new StringBuilder(string.Format("select count(1) from {0} where ", importDataInfo.TableName));
            StringBuilder sbUpdateSql = new StringBuilder(string.Format("update {0} set ", importDataInfo.TableName));
            //先遍历不是多行列的列
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                if (!attributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName))
                {
                    this.fillCountUpdateByAttr(attributeInfo, sbCountSql, sbUpdateSql, ref hasPrimaryKey);
                }
            }

            //再遍历是多行的列
            List<StringBuilder> sbCountList = new List<StringBuilder>();
            List<StringBuilder> sbUpdateList = new List<StringBuilder>();
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                if (attributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName))
                {
                    StringBuilder sbNewCountSql = new StringBuilder(sbCountSql.ToString());
                    StringBuilder sbNewUpdateSql = new StringBuilder(sbUpdateSql.ToString());

                    this.fillCountUpdateByAttr(attributeInfo, sbNewCountSql, sbNewUpdateSql, ref hasPrimaryKey);

                    sbCountList.Add(sbNewCountSql);
                    sbUpdateList.Add(sbNewUpdateSql);
                }
            }    
        
            //移除最后的“， ”和“and ”
            foreach (StringBuilder sbUpdate in sbUpdateList)
            {
                sbUpdate.Remove(sbUpdate.Length - 2, 2);
            }
            if (hasPrimaryKey)
            {
                foreach (StringBuilder sbCount in sbCountList)
                {
                    sbCount.Remove(sbCount.Length - 4, 4);
                }
            }
             
             

            /****先生成没有多列的语句
            bool hasPrimaryKey = false;
            StringBuilder sbCountSql = new StringBuilder(string.Format("select count(1) from {0} where ", importDataInfo.TableName));
            StringBuilder sbUpdateSql = new StringBuilder(string.Format("update {0} set ", importDataInfo.TableName));
            //先遍历不是多行列的列
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                if (!attributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName))
                {
                    this.fillCountUpdateByAttr(attributeInfo, sbCountSql, sbUpdateSql, ref hasPrimaryKey);
                }
            }

            //再遍历是多行的列
            List<StringBuilder> sbCountList = new List<StringBuilder>();
            List<StringBuilder> sbUpdateList = new List<StringBuilder>();
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                if (attributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName))
                {
                    StringBuilder sbNewCountSql = new StringBuilder(sbCountSql.ToString());
                    StringBuilder sbNewUpdateSql = new StringBuilder(sbUpdateSql.ToString());

                    this.fillCountUpdateByAttr(attributeInfo, sbNewCountSql, sbNewUpdateSql, ref hasPrimaryKey);

                    sbCountList.Add(sbNewCountSql);
                    sbUpdateList.Add(sbNewUpdateSql);
                }
            }    
        
            //移除最后的“， ”和“and ”
            foreach (StringBuilder sbUpdate in sbUpdateList)
            {
                sbUpdate.Remove(sbUpdate.Length - 2, 2);
            }
            if (hasPrimaryKey)
            {
                foreach (StringBuilder sbCount in sbCountList)
                {
                    sbCount.Remove(sbCount.Length - 4, 4);
                }
            }*/
#endregion