using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using Newtonsoft.Json;
using System.IO;

namespace TextForCtext
{
    /// <summary>
    /// 使用《看典古籍》OCR API的C#示例
    /// 20240729 Copilot大菩薩：這個 OCRClient 類別包含一個 GetOCRResult 方法，該方法接受一個圖像URL作為參數，並返回OCR結果。請注意，您需要將 "您的token" 和 "oscarsun72@hotmail.com" 替換為您的實際token和註冊郵箱。
    /// </summary>
    internal class OCRClient
    {
        private readonly HttpClient _client;

        public OCRClient()
        {
            _client = new HttpClient();
        }

        public async Task<string> GetOCRResult(string imagePath)
        {
            var bytes = File.ReadAllBytes(imagePath);
            var base64Image = Convert.ToBase64String(bytes);

            var request = new HttpRequestMessage(HttpMethod.Post, "https://ocr.kandianguji.com/ocr_api");

            var json = JsonConvert.SerializeObject(new
            {
                token = "a93a3559-b04a-4b7d-9a83-0338d8cd7681",
                email = "oscarsun72@hotmail.com",
                image = base64Image
            });
            request.Content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _client.SendAsync(request);

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync();
                return result;
            }
            else
            {
                Console.WriteLine($"Error: {response.StatusCode}");
                return null;
            }
        }

        public async Task<string> GetOCRResult_2ndVersion(string imagePath)
        {
            var bytes = File.ReadAllBytes(imagePath);
            var base64Image = Convert.ToBase64String(bytes);

            var request = new HttpRequestMessage(HttpMethod.Post, "https://ocr.kandianguji.com/ocr_api");
            request.Headers.Add("token", "a93a3559-b04a-4b7d-9a83-0338d8cd7681"); // 將 "您的token" 替換為您的實際token
            request.Headers.Add("email", "oscarsun72@hotmail.com"); // 將 "oscarsun72@hotmail.com" 替換為您的實際註冊郵箱

            var json = JsonConvert.SerializeObject(new { image = base64Image });
            request.Content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _client.SendAsync(request);

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync();
                return result;
            }
            else
            {
                Console.WriteLine($"Error: {response.StatusCode}");
                return null;
            }
        }



        internal async Task<string> GetOCRResult_1stVersion(string imageUrl)
        {//public async Task<string> GetOCRResult(string imageUrl)
            var request = new HttpRequestMessage(HttpMethod.Post, "https://www.kandianguji.com/ocr_api");
            //request.Headers.Add("token", "您的token"); // 將 "您的token" 替換為您的實際token
            request.Headers.Add("token", "a93a3559-b04a-4b7d-9a83-0338d8cd7681"); 
            request.Headers.Add("email", "oscarsun72@hotmail.com"); // 將 "oscarsun72@hotmail.com" 替換為您的實際註冊郵箱

            request.Content = new StringContent($"{{\"url\":\"{imageUrl}\"}}", Encoding.UTF8, "application/json");

            var response = await _client.SendAsync(request);

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync();
                return result;
            }
            else
            {
                throw new Exception($"Error: {response.StatusCode}");
            }
        }
    }
}
