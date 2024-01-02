using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Working.Tools.AttributeExpand;

namespace Working.Tools.ModelExpand
{
    public class DataFillModel
    {
        /// <summary>
        /// 命令名称
        /// </summary>        
        public string? CommandName { get; set; }
        /// <summary>
        /// 命令对应的值
        /// </summary>
        public dynamic? Value { get; set; }

        /// <summary>
        /// 行索引
        /// </summary>
        public int RowIndex { get; set; }
        /// <summary>
        /// 列索引
        /// </summary>
        public int CellIndex { get; set; }

        /// <summary>
        /// ValueType 0 是按文字处理，1 是按图片处理
        /// </summary>
        public int ValueType { get; set; }
    }
}
