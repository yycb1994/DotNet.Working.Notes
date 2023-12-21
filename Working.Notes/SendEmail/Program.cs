using System.Data;
using Working.Tools;
using Working.Tools.AttributeExpand;

namespace SendEmail
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            var dt = CreateDataTable();
            dt.ImportExcel("c:\\ImportExcel\\Test.xlsx","导出测试");

            //var columnValidators = new Dictionary<string, Func<object, bool>>
            //  {
            //      { "Age", value => Convert.ToInt32(value) < 31 },
                  
            //       // 添加更多的列和验证函数
            //   };
            //var dt1 = ExcelHelper.ImportExcel("c:\\ImportExcel\\Test.xlsx");
            //var list = CreateDataTable().ToObjectList<Test>();
        }


        static DataTable CreateDataTable()
        {
            // 创建一个 DataTable
            DataTable dataTable = new DataTable("Person");

            // 添加列
            dataTable.Columns.Add("姓名", typeof(string));
            dataTable.Columns.Add("姓名1", typeof(string));
            dataTable.Columns.Add("年龄", typeof(int));

            // 添加行数据
            dataTable.Rows.Add(1,"Alice", 30);
            dataTable.Rows.Add(2,"Bob", 25);
            dataTable.Rows.Add(4,"Charlie", 35);
            dataTable.Rows.Add(3,"David", 20);

            return dataTable;
        }

    }

    public class Test
    {
        [DataTableFieldName("年龄")]
        public int Age { get; set; }
        [DataTableFieldName("姓名")]
        public string Name1 { get; set; }
        [DataTableFieldName("姓名1")]
        public string Name { get; set; }
    }
}
