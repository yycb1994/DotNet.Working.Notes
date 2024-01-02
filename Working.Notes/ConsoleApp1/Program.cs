using RestSharp;
using System.Data;
using Working.Tools;
using Working.Tools.AttributeExpand;

namespace ConsoleApp1
{
    internal class Program
    {
        static void Main(string[] args)
        {

            #region EmailSender Call Example 支持抄送
            //string smtpServer = "smtp.qq.com"; // 或者 "smtp.163.com"  
            //int smtpPort = 587; // 根据SMTP服务器配置更改这个值  网易邮箱端口 25
            //string senderName = "89085824@qq.com"; // 发件人邮箱地址  
            //string senderPassword = "mzodpkjdipvabgibc"; // 发件人邮箱密码  
            //string recipientEmail = "13200813451@163.com"; // 收件人邮箱地址  
            //string cCEmail = "13200813451@163.com"; // 抄送人邮箱地址  
            //string subjectPrefix = $"{DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")}  测试邮件标题";// 邮件主题前缀，可以根据需要更改这个值  
            //string attachmentFilePath = ""; // 附件文件路径，根据需要更改这个值  
            //string emailContext = $"{DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")}  测试邮件内容";

            //EmailSender emailSender = new EmailSender(smtpServer, smtpPort, senderName, senderPassword, subjectPrefix, emailContext);
            //emailSender.SendMail(recipientEmail, cCEmail, attachmentFilePath); // 替换为你需要发送的邮件内容和文件路径，包括附件和图片的路径 
            #endregion

            #region HttpHelper Call Example
            //var requestbody = new RestRequest();
            //var result = HttpHelper.HttpRequest("http://www.baidu.com", RestSharp.Method.Get, requestbody); 
            #endregion

            #region ExcelHelper Call Example
            //var dt = CreateDataTable();
            //dt.ImportExcel("c:\\ImportExcel\\Test.xlsx", "导出测试");

            //var columnValidators = new Dictionary<string, Func<object, bool>>
            //  {
            //      { "Age", value => Convert.ToInt32(value) < 31 },

            //       // 添加更多的列和验证函数
            //   };
            //var dt1 = ExcelHelper.ImportExcel("c:\\ImportExcel\\Test.xlsx");
            //var list = CreateDataTable().ToObjectList<Test>(); 
            #endregion
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
            dataTable.Rows.Add(1, "Alice", 30);
            dataTable.Rows.Add(2, "Bob", 25);
            dataTable.Rows.Add(4, "Charlie", 35);
            dataTable.Rows.Add(3, "David", 20);

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
