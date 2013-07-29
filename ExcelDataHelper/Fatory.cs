using System;
using System.Collections.Generic;
using System.Text;
using ExcelDataHelper.DataManager;

namespace ExcelDataHelper
{
    class Fatory
    {
        private static OracleDataManager oracleDataManager = new OracleDataManager();

        /// <summary>
        /// ��������Ŀ�����͵õ�DataManager(����)
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
