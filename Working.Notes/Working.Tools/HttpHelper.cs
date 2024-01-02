using RestSharp;
using System.Net;

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


        /// <summary>
        /// Http 请求
        /// </summary>
        /// <param name="baseUrl">目标地址</param>
        /// <param name="method">请求类型</param>
        /// <param name="requestBody">body参数</param>
        /// <returns></returns>
        public static string HttpRequest(string baseUrl, Method method, RestRequest requestBody)
        {
            try
            {
                var client = new RestClient(baseUrl);
               
                requestBody.Method = method;
                var response = client.Execute(requestBody);
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    return response.Content;
                }

                throw new Exception("HTTP request result error ." + response.ErrorMessage);
            }
            catch (Exception ex)
            {
                throw new Exception("HTTP request failed ." + ex.Message);
            }

        }
    }
}
