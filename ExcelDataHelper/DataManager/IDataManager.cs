using System;
using System.Collections.Generic;
using System.Text;
using DataImportLib.Entity;
using System.Data;

namespace DataImportLib.DataManager
{
    /// <summary>
    /// 数据管理的抽象接口，主要为了解决不同类型的数据目标（oracle，sqlserver等）之间的不同
    /// </summary>
    interface IDataManager
    {
        /// <summary>
        /// 得到默认的列类型
        /// </summary>
        /// <returns></returns>
        string getDefaultColumnType();


        /// <summary>
        /// 根据数据源表和要导入的列信息导入目的
        /// </summary>
        /// <param name="sourceTable">数据源表，列名跟导入配置sourceName相同的列会被导入</param>
        /// <param name="attributeList"></param>
        /// <param name="importDataInfo"></param>
        /// <param name="containsAllPrimaryKey">是否包含所有的主键，如果是oneToMultiRow则没用，此判断推迟到往数据库添加时</param>
        /// <returns>没有被导入的行</returns>
        DataTable importData(DataTable sourceTable, List<AttributeInfo> attributeList, ImportDataInfo importDataInfo, bool containsAllPrimaryKey);
    }
}
