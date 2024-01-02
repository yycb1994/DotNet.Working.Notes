namespace Working.Tools
{
    /// <summary>
    /// 自定义的定时器
    /// </summary>
    public class CustomizedTimer
    {
        /// <summary>
        /// 创建每日定时任务
        /// </summary>
        /// <param name="action">要执行的任务</param>
        /// <param name="hour">每日几点执行</param>
        /// <param name="minute">每日几分执行 可忽略</param>
        public void CreateDailyScheduledTasks(Action action, int hour, int minute = -1)
        {
            var todayStr = ""; // 存储当天日期的字符串
            bool isRun = false; // 标记当天任务是否已执行过

            while (true)
            {
                DateTime now = DateTime.UtcNow.AddHours(8); // 获取当前时间（北京时间）
                Console.WriteLine($"现在时间[北京]：{now.Year}-{now.Month}-{now.Day} {now.Hour}:{now.Minute}:{now.Second}");

                // 检查当前时间是否满足任务执行条件
                if (now.Hour == hour && (minute < 0 || now.Minute == minute) && !isRun)
                {
                    isRun = true;
                    todayStr = now.ToString("yyyyMMdd"); // 更新当天日期字符串
                    Console.WriteLine($"今天日期：{todayStr}");
                    action.Invoke(); // 执行任务
                    Console.WriteLine("任务已执行");
                }

                // 如果日期发生变化，重置任务执行标记
                if (todayStr != now.ToString("yyyyMMdd"))
                {
                    isRun = false;
                }

                Thread.Sleep(1000); // 可根据实际需求调整延迟时间
            }
        }

    }
}
