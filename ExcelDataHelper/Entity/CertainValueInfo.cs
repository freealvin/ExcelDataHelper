using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataHelper.Entity
{
    /// <summary>
    /// ȷ��ֵ��
    /// </summary>
    public class CertainValueInfo
    {
        private string key;
        private string value;

        /// <summary>
        /// ����Ҳ���������ֵ
        /// </summary>
        public string Key
        {
            get { return key; }
            set { key = value; }
        }

        /// <summary>
        /// ֵ��Ҳ���ǵ��뵽Ŀ���ֵ
        /// </summary>
        public string Value
        {
            get { return this.value; }
            set { this.value = value; }
        }


    }
}
