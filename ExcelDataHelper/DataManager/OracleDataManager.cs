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
    /// Oracle�����ݹ�����
    /// </summary>
    class OracleDataManager : IDataManager
    {
        private static string ParameterChar = ":";
        /// <summary>
        /// �õ�Ĭ�ϵ�������
        /// </summary>
        /// <returns></returns>
        public string getDefaultColumnType()
        {
            return "varchar2";
        }

        /// <summary>
        /// ��������Դ���Ҫ���������Ϣ����Ŀ��
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

        #region ʹ�õ�����еķ�ʽ�������ݿ�
        /// <summary>
        /// ʹ�õ�����еķ�ʽ�������ݿ�
        /// </summary>
        /// <param name="sourceTable"></param>
        /// <param name="attributeList"></param>
        /// <param name="importDataInfo"></param>
        /// <returns></returns>
        private DataTable imoprtDataToMultiRow(DataTable sourceTable, List<AttributeInfo> attributeList, ImportDataInfo importDataInfo)
        {
            /**
             * ����������Insert���
             * insert into tableName(column1, column2, column3......multiColumn, determineColumn1...)
             * select :column1, :column2, :column3......:multiColumn from dual
             * where not exists (select 1 from tableName where primaryKey1 = :primaryKey1 and .....) 
             */

            bool hasPrimaryKey = false;

            #region ��������б�

            StringBuilder sbInsert = new StringBuilder();
            sbInsert.Append(string.Format("insert into {0} (", importDataInfo.TableName));

            //�ȱ������Ƕ����е���
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                if (!attributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName))
                {
                    this.fillSqlByAttrWithDetermineColumn(attributeInfo, "{0}, ", SqlStrType.ColumnName, sbInsert, ref hasPrimaryKey);
                }
            }

            //�ٱ����Ƕ����еģ�����insert����б�
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

            //�Ƴ�����", "���ټ�������һ��
            foreach (StringBuilder sbNewInsert in sbInsertList)
            {
                sbNewInsert.Remove(sbNewInsert.Length - 2, 2);
                sbNewInsert.Append(" ) select ");
            }

            //select�ֶ��б�
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

            //���е��а�˳��������
            int sbInsertIndex = 0;
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                if (attributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName))
                {
                    StringBuilder sbNewInsert = sbInsertList[sbInsertIndex++];
                    this.fillSqlByAttrWithDetermineColumn(attributeInfo, "{0}, ", SqlStrType.ParamName, sbNewInsert, ref hasPrimaryKey);
                }
            }

            //�Ƴ�����", "���ټ�������һ��
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

                //�ж�������where����
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
                //�Ƴ�����"and "���ټ�������������
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

            //��һ����Դ��һ���ı������û�и��³ɹ�������
            DataTable dtBadRecords = CreateTableByAttrList(sourceTable, attributeList);

            try
            {
                conn.Open();

                #region ִ�����
                foreach (DataRow dataRow in sourceTable.Rows)
                {
                    //��Щ���е��еĲ����б�
                    List<OracleParameter> singleRowParamList = new List<OracleParameter>();
                    //�ж�������where����
                    bool hasNullPrimary = false;
                    foreach (AttributeInfo attributeInfo in attributeList)
                    {
                        if (!attributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName))
                        {
                            hasNullPrimary = this.addOracleParameterWithValue(dataRow, attributeInfo, singleRowParamList);
                            if (hasNullPrimary)
                            {
                                //���в��ܱ����룬��Ϊ��������������Ϊ�յ�
                                AddNewRow(dtBadRecords, dataRow);
                                break;
                            }
                        }
                    }

                    if (hasNullPrimary)
                    {
                        //��һ��
                        continue;
                    }

                    int commandTextIndex = 0;
                    //�����еĲ�����һ��ִ��һ��
                    foreach (AttributeInfo attributeInfo in attributeList)
                    {
                        if (attributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName))
                        {
                            comm.CommandText = sbInsertList[commandTextIndex++].ToString();
                            hasNullPrimary = true;
                            //�ȼ���
                            comm.Parameters.Clear();
                            foreach (OracleParameter param in singleRowParamList)
                            {
                                comm.Parameters.Add(param);
                            }

                            hasNullPrimary = this.addOracleParameterWithValue(dataRow, attributeInfo, comm.Parameters);

                            if (hasNullPrimary)
                            {
                                //��һ��
                                continue;
                            }

                            try
                            {
                                comm.ExecuteNonQuery();
                            }
                            catch (OracleException ex)
                            {
                                DataRow newRow = AddNewRow(dtBadRecords, dataRow);

                                //�������������ÿ�
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
                //˵��open�ͳ�����
                throw ex;
            }
            finally
            {
                conn.Close();
            }
            return dtBadRecords;
        }
        /// <summary>
        /// ����OracleParameter���Ҹ���dataRow��ֵ�������ж��Ƿ�����������û��ֵ��
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

            //������ȷ���е��б�
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
        /// ����˷����Ĳ���
        /// </summary>
        enum SqlStrType
        {
            /// <summary>
            /// ֻ��Ҫ����ColumnName
            /// </summary>
            ColumnName,

            /// <summary>
            /// ֻ��Ҫ����ParamName
            /// </summary>
            ParamName,

            /// <summary>
            /// ColumnName��ParamName����Ҫ����
            /// </summary>
            ColumnAndParamName
        }
        /// <summary>
        /// ���ݸ�ʽ�ַ�������attributeInfo����determineColumnInfo����columnName��ӵ�sb��
        /// </summary>
        /// <param name="attributeInfo"></param>
        /// <param name="formatString">ֻ����{0}������ColumnName</param>
        /// <param name="sb"></param>
        private bool fillSqlByAttrWithDetermineColumn(AttributeInfo attributeInfo, string formatString, SqlStrType sqlStrType, StringBuilder sb)
        {
            bool hasPrimaryKey = false;
            this.fillSqlBySqlStrType(attributeInfo, formatString, sqlStrType, sb);
            if (attributeInfo.IsPrimaryKey)
            {
                hasPrimaryKey = true;
            }

            //����DetermineColumnList
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
        /// ����SqlƬ�ε���������sb�����
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
        /// ���ݸ�ʽ�ַ�������attributeInfo����determineColumnInfo����columnName��ӵ�sb��
        /// </summary>
        /// <param name="attributeInfo"></param>
        /// <param name="formatString">ֻ����{0}������ColumnName</param>
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
        /// ����Դ���attributeList����DataTable�����������attributeList��Դ���ﶼ��������
        /// </summary>
        /// <param name="sourceTable"></param>
        /// <param name="attributeList"></param>
        /// <returns></returns>
        private static DataTable CreateTableByAttrList(DataTable sourceTable, List<AttributeInfo> attributeList)
        {
            DataTable newTable = new DataTable();
            //�������attributeList���������
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
        /// ��һ���������һ��
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
        /// ʹ��insert��������
        /// </summary>
        /// <param name="sourceTable"></param>
        /// <param name="attributeList"></param>
        /// <param name="importDataInfo"></param>
        private DataTable importDataByInsert(DataTable sourceTable, List<AttributeInfo> attributeList, ImportDataInfo importDataInfo)
        {
            /**Insert���
             * ʾ����insert into table (column1, column2,....)
             * select :column1, :column2,.... from dual
             * where not exists (select 1 from tableName where primaryKey1 = :primaryKey1 and .....) 
             */

            StringBuilder sbInsert = new StringBuilder("insert into " + importDataInfo.TableName + " ( ");
            //��һ��ѭ��˳��Ѱ����û������
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
                //����������������ж�����
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

            //��һ����Դ��һ���ı������û�и��³ɹ�������
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

                //���в���
                foreach (DataRow dataRow in sourceTable.Rows)
                {
                    try
                    {
                        //��parameter��ֵ
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
                //˵��open�ͳ�����
                throw ex;
            }
            finally
            {
                conn.Close();
            }
            return dtBadRecords;
        }


        /// <summary>
        /// ʹ��Update��������
        /// </summary>
        /// <param name="sourceTable"></param>
        /// <param name="attributeList"></param>
        /// <param name="importDataInfo"></param>
        private DataTable importDataByUpdate(DataTable sourceTable, List<AttributeInfo> attributeList, ImportDataInfo importDataInfo)
        {
            /**
             * Update���(��Ϊ����д����Parameter����Ī�������bug)
             * ʾ���� update table t set column1 = :column1, column2 = :column2....
             *  where t.primaryKey1 = :primaryKey1 and t.primaryKey2 = :primaryKey2...
             */
            List<AttributeInfo> primaryKeyList = importDataInfo.PrimaryKeyList;
            if (primaryKeyList.Count == 0)
            {
                throw new Exception("ʹ�ø��·���������������");
            }

            StringBuilder sbMerge = new StringBuilder(string.Format(" update {0} t set ", importDataInfo.TableName));
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                //ֻupdate������������
                if (!attributeInfo.IsPrimaryKey)
                {
                    //"columnName = :columnName, "
                    sbMerge.Append(string.Format("{0} = {1}, ", attributeInfo.ColumnName, GetAttributeParameterSql(attributeInfo)));
                }
            }
            sbMerge.Remove(sbMerge.Length - 2, 2);

            sbMerge.Append(" where ");

            //where����
            foreach (AttributeInfo attributeInfo in primaryKeyList)
            {
                //"t2.columnName = t.columnName and "
                sbMerge.Append(string.Format("t.{0} = {1} and ", attributeInfo.ColumnName, GetAttributeParameterSql(attributeInfo)));
            }
            sbMerge.Remove(sbMerge.Length - 4, 4);


            OracleConnection conn = new OracleConnection(GetConnectionString(importDataInfo));
            OracleCommand comm = new OracleCommand(sbMerge.ToString(), conn);

            //��һ����Դ��һ���ı������û�и��³ɹ�������
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
                //���в���
                foreach (DataRow dataRow in sourceTable.Rows)
                {
                    try
                    {
                        //��parameter��ֵ
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
                        //���û��Ӱ�����˵��Ҳû�гɹ�
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
                //˵��open�ͳ�����
                throw ex;
            }
            finally
            {
                conn.Close();
            }
            return dtBadRecords;
        }


        /// <summary>
        /// ʹ��merge��������
        /// </summary>
        /// <param name="sourceTable"></param>
        /// <param name="attributeList"></param>
        /// <param name="importDataInfo"></param>
        private DataTable importDataByMerge(DataTable sourceTable, List<AttributeInfo> attributeList, ImportDataInfo importDataInfo)
        {
            /**
             * Merge���(��Ϊ����д����Parameter����Ī�������bug)
             * ʾ����merge into table t
             * using ( select '1' c from dual ) t2 on (t2.c = '1' and t.primaryKey1 = :primaryKey1 and t.primaryKey2 = :primaryKey2...)
             * when matched then update set column1 = :column1, column2 = :column2....
             * when not matched then insert (column1, column2, .....) values (:column1, :column2,.......)
             */

            StringBuilder sbMerge = new StringBuilder(string.Format(" merge into {0} t ", importDataInfo.TableName));
            sbMerge.Append(" using (select '1' c from dual) t2 on (t2.c = '1' and ");
            List<AttributeInfo> primaryKeyList = importDataInfo.PrimaryKeyList;
            //t��on����
            foreach (AttributeInfo attributeInfo in primaryKeyList)
            {
                //"t2.columnName = t.columnName and "
                sbMerge.Append(string.Format("t.{0} = {1}{0} and ", attributeInfo.ColumnName, ParameterChar));
            }
            sbMerge.Remove(sbMerge.Length - 4, 4);

            //update��set�ֶ��б�
            sbMerge.Append(" ) when matched then update set ");
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                //ֻupdate������������
                if (!attributeInfo.IsPrimaryKey)
                {
                    //"columnName = :columnName, "
                    sbMerge.Append(string.Format("{0} = {1}, ", attributeInfo.ColumnName, GetAttributeParameterSql(attributeInfo)));
                }
            }
            sbMerge.Remove(sbMerge.Length - 2, 2);

            //insert�ֶ�
            sbMerge.Append(" when not matched then insert (");
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                sbMerge.Append(attributeInfo.ColumnName + ", ");
            }
            sbMerge.Remove(sbMerge.Length - 2, 2);

            //insertֵ�ֶ�
            sbMerge.Append(" ) values (");
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                sbMerge.Append(GetAttributeParameterSql(attributeInfo) + ", ");
            }
            sbMerge.Remove(sbMerge.Length - 2, 2);
            sbMerge.Append(" )");


            OracleConnection conn = new OracleConnection(GetConnectionString(importDataInfo));
            OracleCommand comm = new OracleCommand(sbMerge.ToString(), conn);

            //��һ����Դ��һ���ı������û�и��³ɹ�������
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
                //���в���
                foreach (DataRow dataRow in sourceTable.Rows)
                {
                    try
                    {
                        //��parameter��ֵ
                        int paraPos = 0;
                        foreach (AttributeInfo primaryAttributeInfo in primaryKeyList)
                        {
                            comm.Parameters[paraPos++].Value = GetAttributeParameterValue(dataRow, primaryAttributeInfo);
                        }

                        foreach (AttributeInfo attributeInfo in attributeList)
                        {
                            if (!attributeInfo.IsPrimaryKey)
                            {
                                if (attributeInfo.SourceName.StartsWith("ѧʿ"))
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
                //˵��open�ͳ�����
                throw ex;
            }
            finally
            {
                conn.Close();
            }
            return dtBadRecords;
        }

        /// <summary>
        /// �õ�����Parameter������
        /// </summary>
        /// <param name="attributeInfo"></param>
        /// <returns></returns>
        private static SqlDbType GetAttributeDbType(BaseAttributeInfo baseAttributeInfo)
        {
            SqlDbType dbType = (SqlDbType)Enum.Parse(typeof(SqlDbType), baseAttributeInfo.ValueType, true);
            return dbType;
        }

        /// <summary>
        /// �õ����е�Parameter��ʾsql,һ����:columnName
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
        /// �õ����е�Parameter��ֵ,���Ϊnull����ַ��������ַ�������
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

                //��ǰ��ո�ȥ��
                if (attributeInfo.Trim)
                {
                    strData = strData.Trim();
                    if (string.IsNullOrEmpty(strData))
                    {
                        ////ȥ���ո��Ϊ���ַ�����������ַ����ͣ��򷵻��ַ�����������ǣ��򷵻ؿ�
                        //switch (attributeInfo.AttributeType.ToLower())
                        //{
                        //    case "varchar2":
                        //        return strData;
                        //    default:
                        //        return DBNull.Value;
                        //}

                        //��oracle����ַ�������null������
                        return DBNull.Value;
                    }
                }
                //ȡ��ȷ��ֵ
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
                        throw new Exception(attributeInfo.SourceName + "��ֵ���ڹ涨��ֵ�У�");
                    }
                }
                //�������͵�ת��DateTime
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

                //�������ͷ���
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
                throw new Exception(attributeInfo.SourceName + "����" + ex.Message, ex);
            }
        }

        /// <summary>
        /// �õ����е�Parameter��ֵ
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
                    //��oracle����ַ�������null������
                    return DBNull.Value;
                }

                //�������͵�ת��DateTime
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

                //�������ͷ���
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
                throw new Exception(determineColumnInfo.ColumnName + "����" + ex.Message, ex);
            }
        }
       
        /// <summary>
        /// �õ������ַ���
        /// </summary>
        /// <param name="importDataInfo"></param>
        /// <returns></returns>
        private static string GetConnectionString(ImportDataInfo importDataInfo)
        {
            return string.IsNullOrEmpty(importDataInfo.ConnectionString) ? Globals.DefaultConnectionString : importDataInfo.ConnectionString;
        }
    }
}

#region �ϴ���
/****������û�ж��е����
            bool hasPrimaryKey = false;
            StringBuilder sbCountSql = new StringBuilder(string.Format("select count(1) from {0} where ", importDataInfo.TableName));
            StringBuilder sbUpdateSql = new StringBuilder(string.Format("update {0} set ", importDataInfo.TableName));
            //�ȱ������Ƕ����е���
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                if (!attributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName))
                {
                    this.fillCountUpdateByAttr(attributeInfo, sbCountSql, sbUpdateSql, ref hasPrimaryKey);
                }
            }

            //�ٱ����Ƕ��е���
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
        
            //�Ƴ����ġ��� ���͡�and ��
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
             
             

            /****������û�ж��е����
            bool hasPrimaryKey = false;
            StringBuilder sbCountSql = new StringBuilder(string.Format("select count(1) from {0} where ", importDataInfo.TableName));
            StringBuilder sbUpdateSql = new StringBuilder(string.Format("update {0} set ", importDataInfo.TableName));
            //�ȱ������Ƕ����е���
            foreach (AttributeInfo attributeInfo in attributeList)
            {
                if (!attributeInfo.ColumnName.Equals(importDataInfo.MultiRowColumnName))
                {
                    this.fillCountUpdateByAttr(attributeInfo, sbCountSql, sbUpdateSql, ref hasPrimaryKey);
                }
            }

            //�ٱ����Ƕ��е���
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
        
            //�Ƴ����ġ��� ���͡�and ��
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