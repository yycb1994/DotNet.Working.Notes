using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Working.Tools
{
    public class FileHelper
    {
        /// <summary>
        /// 将内容写入到指定的文件中
        /// </summary>
        /// <param name="directoryPath">文件存放的目录路径</param>
        /// <param name="fileName">文件名称</param>
        /// <param name="content">要写入的内容</param>
        /// <returns>是否成功写入文件</returns>
        public static bool WriteToFile(string directoryPath, string fileName, string content)
        {
            try
            {
                string filePath = Path.Combine(directoryPath, fileName); // 组合目录路径和文件名称
                string directory = Path.GetDirectoryName(filePath); // 获取目录路径
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory); // 创建目录
                }

                File.WriteAllText(filePath, content); // 写入文件内容
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"写入文件时出现错误：{ex.Message}");
                return false;
            }
        }
    }
}
