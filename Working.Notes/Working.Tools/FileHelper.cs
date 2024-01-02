using System.IO.Compression;

namespace Working.Tools
{
    public class FileHelper
    {
        #region 将内容写入到指定的文件中

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
        #endregion

        #region 根据文件的url将文件下载到本地
        /// <summary>
        /// 根据文件的url将文件下载到本地
        /// </summary>
        /// <param name="downLoadUrl"></param>
        /// <param name="savePath"></param>
        /// <param name="fileExtension"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static async Task<string> FileDownLoad(string downLoadUrl, string savePath, string fileExtension, string fileName = "")
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // 发送GET请求并获取数据
                    byte[] fileData = await client.GetByteArrayAsync(downLoadUrl);

                    // 保存到本地文件
                    fileName = string.IsNullOrEmpty(fileName) ? CreateFileName(fileExtension) : CreateFileName(fileName, fileExtension);
                    var filepath = Path.Combine(savePath, fileName);
                    File.WriteAllBytes(filepath, fileData);
                    return filepath;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"文件下载错误{ex.Message}");
            }
        }
        #endregion

        #region 生成一个随机数字的文件名称      
        /// <summary>
        /// 生成一个随机数字的文件名称      
        /// </summary>
        /// <param name="extension">后缀名（无须加.）</param>
        /// <returns>文件名称.后缀名</returns>
        public static string CreateFileName(string extension)
        {
            return $"{SnowflakeIdGenerator.CreateNextId()}.{extension}";
        }
        #endregion

        #region 生成一个指定的文件名称
        /// <summary>
        /// 生成一个指定的文件名称
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="extension">后缀名（无须加.）</param>
        /// <returns>文件名称.后缀名</returns>
        public static string CreateFileName(string fileName, string extension)
        {
            return $"{fileName}.{extension}";
        } 
        #endregion

        #region 将一个文件转换为byte[]
        /// <summary>
        /// 将一个文件转换为byte[]
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>byte[]</returns>
        public static byte[] FileConvertToByteArray(string filePath)
        {
            if (File.Exists(filePath))
            {
                return File.ReadAllBytes(filePath);
            }
            return null;
        }
        #endregion

        #region 将指定目录创建为压缩文件
        /// <summary>
        /// 将指定目录创建为压缩文件
        /// </summary>
        /// <param name="sourceFolderPath">要压缩的文件夹路径</param>
        /// <param name="destinationZipPath">生成压缩文件的保存路径</param>
        public static void CreateZipFile(string sourceFolderPath, string destinationZipPath)
        {
            ZipFile.CreateFromDirectory(sourceFolderPath, destinationZipPath);
            //.net4   FastZip fastZip = new FastZip();
            //        fastZip.CreateZip(destinationZipPath, sourceFolderPath, true, "");
        }
        #endregion

        #region 删除指定目录下的所有文件夹和文件
        /// <summary>
        /// 删除指定目录下的所有文件夹和文件
        /// </summary>
        /// <param name="folderPath">指定目录路径</param>
        public static void DeleteFolderContents(string folderPath)
        {
            foreach (string file in Directory.GetFiles(folderPath))
            {
                File.Delete(file);
            }

            foreach (string subfolder in Directory.GetDirectories(folderPath))
            {
                DeleteFolderContents(subfolder);
            }

            Directory.Delete(folderPath, true);
        }
        #endregion

        #region 根据输入的文件夹路径创建文件目录
        /// <summary>
        /// 根据输入的文件夹路径创建文件目录
        /// </summary>
        /// <param name="directoryPath">文件夹路径</param>
        /// <param name="isDelFile">如果该路径下的文件夹存在，是否清空文件夹中的文件</param>
        /// <returns>文件夹路径</returns>
        public static string CreateDirectoryPath(string directoryPath, bool isDelFile = false)
        {
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }

            if (isDelFile)
            {
                DeleteFolderContents(directoryPath);
            }
            return directoryPath;
        }
        #endregion

        /// <summary>
        /// 将给定的字节数组保存为指定路径下的文件。
        /// </summary>
        /// <param name="bytes">要保存为文件的字节数组。</param>
        /// <param name="saveFileFullPath">要保存文件的完整路径。</param>
        public static void SaveFile(byte[] bytes, string saveFileFullPath)
        {
            File.WriteAllBytes(saveFileFullPath, bytes);
        }

    }
}
