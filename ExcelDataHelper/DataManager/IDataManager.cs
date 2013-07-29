using System;
using System.Collections.Generic;
using System.Text;
using DataImportLib.Entity;
using System.Data;

namespace DataImportLib.DataManager
{
    /// <summary>
    /// ���ݹ���ĳ���ӿڣ���ҪΪ�˽����ͬ���͵�����Ŀ�꣨oracle��sqlserver�ȣ�֮��Ĳ�ͬ
    /// </summary>
    interface IDataManager
    {
        /// <summary>
        /// �õ�Ĭ�ϵ�������
        /// </summary>
        /// <returns></returns>
        string getDefaultColumnType();


        /// <summary>
        /// ��������Դ���Ҫ���������Ϣ����Ŀ��
        /// </summary>
        /// <param name="sourceTable">����Դ����������������sourceName��ͬ���лᱻ����</param>
        /// <param name="attributeList"></param>
        /// <param name="importDataInfo"></param>
        /// <param name="containsAllPrimaryKey">�Ƿ�������е������������oneToMultiRow��û�ã����ж��Ƴٵ������ݿ����ʱ</param>
        /// <returns>û�б��������</returns>
        DataTable importData(DataTable sourceTable, List<AttributeInfo> attributeList, ImportDataInfo importDataInfo, bool containsAllPrimaryKey);
    }
}
