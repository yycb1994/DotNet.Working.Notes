using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Working.Tools
{
    public class HttpHelper
    {

        /// <summary>
        /// 从指定的 URL 获取文件内容
        /// </summary>
        /// <param name="url">要请求的 URL</param>
        /// <param name="method">HTTP 请求方法</param>
        /// <param name="postData">POST 请求的数据（可选）</param>
        /// <returns>文件内容</returns>
        public static async Task<string> GetFileContent(string url, HttpMethod method, string postData = null)
        {
            using (HttpClient client = new HttpClient())
            {
                HttpRequestMessage request = new HttpRequestMessage(method, url);
                if (method == HttpMethod.Post && !string.IsNullOrEmpty(postData))
                {
                    request.Content = new StringContent(postData);
                }

                HttpResponseMessage response = await client.SendAsync(request);
                if (response.IsSuccessStatusCode)
                {
                    return await response.Content.ReadAsStringAsync();
                }
                else
                {
                    Console.WriteLine($"请求失败：{response.StatusCode}");
                    return null;
                }
            }
        }
    }
}
