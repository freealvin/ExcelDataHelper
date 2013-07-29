using System;
using System.Collections.Generic;
using System.Text;
using DataImportLib.DataManager;

namespace DataImportLib
{
    class Fatory
    {
        private static OracleDataManager oracleDataManager = new OracleDataManager();

        /// <summary>
        /// 根据数据目标类型得到DataManager(单例)
        /// </summary>
        /// <param name="destinationType"></param>
        /// <returns></returns>
        public static IDataManager GetDataManagerByType(DestinationType destinationType)
        {
            switch (destinationType)
            {
                case DestinationType.Oracle:
                    return oracleDataManager;
                default:
                    return oracleDataManager;
            }
        }
    }
}
