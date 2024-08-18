using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using Newtonsoft.Json;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Threading;
using System.Net.Http.Headers;
using WindowsFormsApp1;
using System.Windows.Forms;
using System.Net;

namespace TextForCtext
{
    /// <summary>
    /// 使用《看典古籍》OCR API的C#示例
    /// 20240729 Copilot大菩薩：這個 OCRClient 類別包含一個 GetOCRResult 方法，該方法接受一個圖像URL作為參數，並返回OCR結果。請注意，您需要將 "您的token" 和 "oscarsun72@hotmail.com" 替換為您的實際token和註冊郵箱。
    /// </summary>
    internal class OCRClient
    {
        private readonly HttpClient _client;
        public Form1 ActiveForm1;

        public OCRClient()
        {
            _client = new HttpClient();
            ActiveForm1 = Application.OpenForms[0] as Form1;
        }

        /// <summary>
        /// 取得《看典古籍》OCR API 執行的結果
        /// </summary>
        /// <param name="imagePath">本機圖檔全檔名</param>
        /// <returns>回傳執行結果的字串</returns>
        public string GetOCRResult(string imagePath)
        {//20240731 Copilot大菩薩：解決圖檔讀取問題的程式碼修改建議：看來這個錯誤是因為圖檔還未完全下載完成就被讀取了。您可以在讀取圖檔之前加入一個等待機制，確保圖檔已經完全下載。以下是修改後的程式碼：
         //int retryCntr = 0;
         //retry:
         // 等待圖檔完全下載            
            while (!File.Exists(imagePath))
            {
                Thread.Sleep(100); // 等待 100 毫秒
            }
            ActiveForm1.TopMost = false;
            // 確保圖檔可以被讀取
            bool fileReady = false;
            while (!fileReady)
            {
                try
                {
                    using (FileStream stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                    {
                        fileReady = true;
                    }
                }
                catch (IOException)
                {
                    Thread.Sleep(100); // 等待 100 毫秒
                }
            }

            var bytes = File.ReadAllBytes(imagePath);
            var base64Image = Convert.ToBase64String(bytes);

            var request = new HttpRequestMessage(HttpMethod.Post, "https://ocr.kandianguji.com/ocr_api");//https://images.kandianguji.com:14141/ocr_api

            // 讀取 token 和 email
            string tokenPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "CtextTempFiles", "OCRAPItoken.txt");
            string token = File.ReadAllText(tokenPath).Trim();

            var json = JsonConvert.SerializeObject(new
            {//https://kandianguji.com/ocr_api_doc
                token = token,
                email = "oscarsun72@hotmail.com",
                image = base64Image,
                version = "beta",//這是站長在微信跟我說的，相當於網頁版的「语序优化beta版」選項 https://kandianguji.com/ocr
                det_mode = "sp"//det_mode：文字内容排版样式，目前有三种可选：auto（自动识别）、sp（竖向排版）、hp（横向排版）；字符串类型，默认值：auto
            });
            request.Content = new StringContent(json, Encoding.UTF8, "application/json");

            HttpResponseMessage response = null;

            //20240818:Copilot大菩薩：.NET Framework 版本之間的 SSL/TLS 支援差異:https://sl.bing.net/fV3va6ogkqi
            //為什麼在執行這個行時，以 .NET Framework 4.7.2 建置的程式不會出錯，但用 .NET Framework 4.5.2 建置的卻會出現「WebException: 要求已經中止: 無法建立 SSL/TLS 的安全通道」的錯誤呢（詳後所附）
            //  HResult=0x80131500
            //  Message = 傳送要求時發生錯誤。
            //response = _client.SendAsync(request).GetAwaiter().GetResult();
            /*這個問題可能是由於 .NET Framework 版本之間的 SSL/TLS 支援差異所引起的。
                .NET Framework 4.5.2 預設使用 TLS 1.0 和 TLS 1.1，這些版本已經被認為是不安全的12。而 .NET Framework 4.7.2 預設使用 TLS 1.2，並且允許使用 TLS 1.1 和 TLS 1.012。這意味著在 .NET Framework 4.7.2 中，應用程式可以自動使用更安全的 TLS 版本，從而避免了「無法建立 SSL/TLS 的安全通道」的錯誤。
                要解決這個問題，您可以在 .NET Framework 4.5.2 中明確指定使用 TLS 1.2。您可以在程式碼中加入以下設定：
             */
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            // 這樣，您的應用程式將會使用 TLS 1.2 來建立安全通道。

            try
            {//20240804 Copilot大菩薩： 使用 CancellationToken 設置程式超時機制：  https://sl.bing.net/jBHhw5kprxI
             //您可以使用 CancellationToken 來設置一個超時機制，避免程式在 SendAsync 這一行卡住。以下是修改後的程式碼：
             //這樣，當請求超過 10 秒未完成時，會自動取消請求並進行重試。希望這樣的修改能夠解決您的問題！如果還有其他問題，隨時告訴我哦！
             //using (var cts = new CancellationTokenSource(TimeSpan.FromSeconds(10))) // 設置 10 秒超時
             //using (var cts = new CancellationTokenSource(TimeSpan.FromSeconds(40))) // 設置 40 秒超時
             //{
             //因為重送講求會被扣點數，先取消

                //response = _client.SendAsync(request, cts.Token).GetAwaiter().GetResult();
                response = _client.SendAsync(request).GetAwaiter().GetResult();
                //}
            }
            catch (OperationCanceledException)
            {
                //Console.WriteLine("Request timed out.");
                //if (retryCntr < 3)
                //{
                //    Thread.Sleep(1500);
                //    retryCntr++;
                //    Form1.playSound(Form1.soundLike.processing, true);                    
                //    goto retry;
                //}
                //else
                return null;
            }
            //catch (Exception ex)
            catch (Exception)
            {
                //Console.WriteLine($"Error: {ex.Message}");
                //if (retryCntr < 3)
                //{
                //    Thread.Sleep(1500);
                //    retryCntr++;
                //    Form1.playSound(Form1.soundLike.processing, true);
                //    goto retry;
                //}
                //else
                return null;
            }
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
