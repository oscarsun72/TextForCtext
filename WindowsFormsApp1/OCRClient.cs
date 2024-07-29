using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using Newtonsoft.Json;
using System.IO;
using Newtonsoft.Json.Linq;

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


        public string GetOCRResult(string imagePath)
        {
            var bytes = File.ReadAllBytes(imagePath);
            var base64Image = Convert.ToBase64String(bytes);

            var request = new HttpRequestMessage(HttpMethod.Post, "https://ocr.kandianguji.com/ocr_api");

            // 讀取 token 和 email
            string tokenPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "CtextTempFiles", "OCRAPItoken.txt");
            string token = File.ReadAllText(tokenPath).Trim();

            var json = JsonConvert.SerializeObject(new
            {
                token = token,
                email = "oscarsun72@hotmail.com",
                image = base64Image
            });
            request.Content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = _client.SendAsync(request).GetAwaiter().GetResult();

            if (response.IsSuccessStatusCode)
            {
                var result = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();

                // 解析 JSON
                var parsedJson = JObject.Parse(result);
                var data = parsedJson["data"].ToObject<List<string>>();

                // 將 data 轉換為換行分隔的字符串
                var formattedResult = string.Join(Environment.NewLine, data);

                return formattedResult;
            }
            else
            {
                Console.WriteLine($"Error: {response.StatusCode}");
                return null;
            }
        }
        /// <summary>
        /// 20240730 Copilot大菩薩： 在這個版本中，我們首先從 OCRAPItoken.txt 檔案中讀取 token，然後將其用於 API 請求。請注意，您需要將 OCRAPItoken.txt 檔案放在 C:\Users\oscar\Documents\CtextTempFiles 目錄下，並將您的 token 寫入該檔案。
        /// </summary>
        /// <param name="imagePath"></param>
        /// <returns></returns>
        public async Task<string> GetOCRResult_5thVersion(string imagePath)
        {
            var bytes = File.ReadAllBytes(imagePath);
            var base64Image = Convert.ToBase64String(bytes);

            var request = new HttpRequestMessage(HttpMethod.Post, "https://ocr.kandianguji.com/ocr_api");

            // 讀取 token 和 email
            string tokenPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "CtextTempFiles", "OCRAPItoken.txt");
            string token = File.ReadAllText(tokenPath).Trim();

            var json = JsonConvert.SerializeObject(new
            {
                token = token,
                email = "oscarsun72@hotmail.com",
                image = base64Image
            });
            request.Content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _client.SendAsync(request);


            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync();

                // 解析 JSON
                var parsedJson = JObject.Parse(result);
                var data = parsedJson["data"].ToObject<List<string>>();

                // 將 data 轉換為換行分隔的字符串
                var formattedResult = string.Join(Environment.NewLine, data);

                return formattedResult;
            }
            else
            {
                Console.WriteLine($"Error: {response.StatusCode}");
                return null;
            }
        }


        /// <summary>
        /// 20240730 Copilot大菩薩：使用《看典古籍》OCR API的C#示例
        /// 在這個版本中，我們首先解析回傳的JSON，然後從 data 鍵中獲取文字列表。然後，我們將這個列表轉換為一個由換行符分隔的字符串，這樣每個文字都會在新的一行。
        ///請注意，這個方法假設回傳的JSON總是包含一個名為 data 的鍵，並且其值總是一個字符串列表。如果API的回傳格式有所變化，您可能需要更新這個方法以適應新的格式。
        /// </summary>
        /// <param name="imagePath"></param>
        /// <returns></returns>
        public async Task<string> GetOCRResult_4thVersion(string imagePath)
        {
            var bytes = File.ReadAllBytes(imagePath);
            var base64Image = Convert.ToBase64String(bytes);

            var request = new HttpRequestMessage(HttpMethod.Post, "https://ocr.kandianguji.com/ocr_api");

            var json = JsonConvert.SerializeObject(new
            {
                token = "***",
                email = "oscarsun72@hotmail.com",
                image = base64Image
            });
            request.Content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _client.SendAsync(request);

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync();

                // 解析 JSON
                var parsedJson = JObject.Parse(result);
                var data = parsedJson["data"].ToObject<List<string>>();

                // 將 data 轉換為換行分隔的字符串
                var formattedResult = string.Join(Environment.NewLine, data);

                return formattedResult;
            }
            else
            {
                Console.WriteLine($"Error: {response.StatusCode}");
                return null;
            }
        }

        public async Task<string> GetOCRResult_3thVersion(string imagePath)
        {
            var bytes = File.ReadAllBytes(imagePath);
            var base64Image = Convert.ToBase64String(bytes);

            var request = new HttpRequestMessage(HttpMethod.Post, "https://ocr.kandianguji.com/ocr_api");

            var json = JsonConvert.SerializeObject(new
            {
                token = "***",
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
            request.Headers.Add("token", "***"); // 將 "您的token" 替換為您的實際token
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
            request.Headers.Add("token", "***"); 
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
