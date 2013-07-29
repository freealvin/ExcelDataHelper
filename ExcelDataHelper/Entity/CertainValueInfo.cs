using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataHelper.Entity
{
    /// <summary>
    /// 确定值类
    /// </summary>
    public class CertainValueInfo
    {
        private string key;
        private string value;

        /// <summary>
        /// 键，也就是输入的值
        /// </summary>
        public string Key
        {
            get { return key; }
            set { key = value; }
        }

        /// <summary>
        /// 值，也就是导入到目标的值
        /// </summary>
        public string Value
        {
            get { return this.value; }
            set { this.value = value; }
        }


    }
}
