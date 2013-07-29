using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace ExcelDataHelper.Entity
{
    /// <summary>
    /// 导入结果类
    /// </summary>
    public class ImportResult
    {
        private string typeName;
        private DataTable dtBadRecords;
        private bool allFailed;

        /// <summary>
        /// 导入的类型名称
        /// </summary>
        public string TypeName
        {
            get { return this.typeName; }
            set { this.typeName = value; }
        }

        /// <summary>
        /// 没导入的记录
        /// </summary>
        public DataTable DtBadRecords
        {
            get { return this.dtBadRecords; }
            set { this.dtBadRecords = value; }
        }

        /// <summary>
        /// 全都没导入
        /// </summary>
        public bool AllFailed
        {
            get { return allFailed; }
            set { allFailed = value; }
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="typeName"></param>
        /// <param name="dtBadRecords"></param>
        public ImportResult(string typeName, DataTable dtBadRecords, bool allFailed)
        {
            this.typeName = typeName;
            this.dtBadRecords = dtBadRecords;
            this.allFailed = allFailed;
        }
    }
}
