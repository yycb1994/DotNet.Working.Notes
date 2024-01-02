namespace Working.Tools.AttributeExpand
{
    public class DataTableFieldNameAttribute: Attribute
    {
        // 列名
        public string ColumnName { get; }

        // 构造函数，用于设置列名
        public DataTableFieldNameAttribute(string columnName)
        {
            ColumnName = columnName;
        }
    }
}
