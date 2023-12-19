using Working.Tools;

namespace SendEmail
{
    internal class Program
    {
        static void Main(string[] args)
        {
            CustomizedTimer  customizedTimer = new CustomizedTimer();
            customizedTimer.CreateDailyScheduledTasks(() => { Console.WriteLine("你好"); },18);

        }
    }
}
