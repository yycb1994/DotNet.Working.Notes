using Working.Tools;

namespace SendEmail
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            //CustomizedTimer  customizedTimer = new CustomizedTimer();
            //customizedTimer.CreateDailyScheduledTasks(() => { Console.WriteLine("你好"); },18);

            var d = await HttpHelper.GetFileContent("http://www.baidu.com", HttpMethod.Get);
        }
    }
}
