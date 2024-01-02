using System.Diagnostics;

namespace Working.Tools
{
    /// <summary>
    /// 进程执行器【通过进程启动信息对象执行后台命令】
    /// </summary>
    public static class ProcessExecutor
    {
        /// <summary>
        /// 执行某个脚本并返回结果
        /// </summary>
        /// <param name="executable">可执行文件路径</param>
        /// <param name="arguments">要传递给脚本的命令行参数</param>
        /// <returns>结果</returns>
        public static string ExecuteScript(string executable, string arguments)
        {
            // 创建一个新的进程启动信息对象
            ProcessStartInfo startInfo = new ProcessStartInfo();

            // 设置要执行的可执行文件路径
            startInfo.FileName = executable;

            // 设置要传递给脚本的命令行参数
            startInfo.Arguments = arguments;

            // 禁用使用操作系统外壳启动进程
            startInfo.UseShellExecute = false;

            // 将标准输出重定向到StreamReader以便读取脚本的输出
            startInfo.RedirectStandardOutput = true;

            // 使用进程启动信息创建新的进程对象
            using (Process process = Process.Start(startInfo))
            {
                // 创建一个StreamReader以读取进程的标准输出
                using (StreamReader reader = process.StandardOutput)
                {
                    // 读取进程的标准输出，即脚本的输出结果
                    string result = reader.ReadToEnd();

                    // 返回脚本的输出结果
                    return result;
                }
            }
        }

    }
}
