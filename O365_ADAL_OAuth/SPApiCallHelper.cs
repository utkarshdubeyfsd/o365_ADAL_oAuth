using Microsoft.Graph;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Syroot.Windows.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Directory = System.IO.Directory;
using File = System.IO.File;

namespace O365_ADAL_OAuth
{
    /// <summary>
    /// Helper class to call a SharePoint graph API and process its result
    /// </summary>
    public class SPApiCallHelper
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="httpClient">HttpClient used to call the SharePoint graph API</param>
        public SPApiCallHelper(HttpClient httpClient)
        {
            HttpClient = httpClient;
        }

        protected HttpClient HttpClient { get; private set; }

        /// <summary>
        /// upload file to sharepoint site, using sharepoint graph API
        /// </summary>
        /// <param name="webApiUrl">Url of the Web API to call (supposed to return Json)</param>
        /// <param name="accessToken">Access token used as a bearer security token to call the Web API</param>
        /// <param name="filePath">path of file, which is going to upload on sharepoint</param>
        /// <param name="processResult">Callback used to process the result of the call to the Web API</param>
        /// <returns></returns>
        public async Task UploadFile(string webApiUrl, string accessToken, string filePath, Action<JObject> processResult)
        {
            if (!string.IsNullOrEmpty(accessToken))
            {
                var defaultRequetHeaders = HttpClient.DefaultRequestHeaders;
                if (defaultRequetHeaders.Accept == null || !defaultRequetHeaders.Accept.Any(m => m.MediaType == "application/json"))
                {
                    HttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                }
                defaultRequetHeaders.Authorization = new AuthenticationHeaderValue("bearer", accessToken);

                HttpResponseMessage response = await HttpClient.PutAsync(webApiUrl, new ByteArrayContent(File.ReadAllBytes(filePath)));
                if (response.IsSuccessStatusCode)
                {
                    string json = await response.Content.ReadAsStringAsync();
                    JObject result = JsonConvert.DeserializeObject(json) as JObject;
                    Console.ForegroundColor = ConsoleColor.Gray;
                    processResult(result);
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Failed to call the Web Api: {response.StatusCode}");
                    string content = await response.Content.ReadAsStringAsync();

                    // Note that if you got reponse.Code == 403 and reponse.content.code == "Authorization_RequestDenied"
                    // this is because the tenant admin as not granted consent for the application to call the Web API
                    Console.WriteLine($"Content: {content}");
                }
                Console.ResetColor();
            }
        }

        /// <summary>
        /// download file from sharepoint site, using sharepoint graph API
        /// </summary>
        /// <param name="webApiUrl">Url of the Web API to call (supposed to return Json)</param>
        /// <param name="accessToken">Access token used as a bearer security token to call the Web API</param>
        /// <returns></returns>
        public async Task DownloadFile(string webApiUrl, string accessToken)
        {
            string downloadsPath = new KnownFolder(KnownFolderType.Downloads).Path;

            if (!string.IsNullOrEmpty(accessToken))
            {
                var defaultRequetHeaders = HttpClient.DefaultRequestHeaders;
                if (defaultRequetHeaders.Accept == null || !defaultRequetHeaders.Accept.Any(m => m.MediaType == "application/json"))
                {
                    HttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                }
                defaultRequetHeaders.Authorization = new AuthenticationHeaderValue("bearer", accessToken);

                HttpResponseMessage response = await HttpClient.GetAsync(webApiUrl);
                if (response.IsSuccessStatusCode)
                {
                    using (Stream contentStream = await response.Content.ReadAsStreamAsync(),
                    stream = new FileStream(downloadsPath + "\\" + response.Content.Headers.ContentDisposition.FileNameStar, FileMode.Create, FileAccess.Write))
                    {
                        await contentStream.CopyToAsync(stream);
                    }

                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("File Downloaded at location: " + downloadsPath);

                    Console.ResetColor();
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Failed to call the Web Api: {response.StatusCode}");
                    string content = await response.Content.ReadAsStringAsync();

                    // Note that if you got reponse.Code == 403 and reponse.content.code == "Authorization_RequestDenied"
                    // this is because the tenant admin as not granted consent for the application to call the Web API
                    Console.WriteLine($"Content: {content}");
                }
                Console.ResetColor();
            }
        }

        /// <summary>
        /// copy file between sharepoint libraries, using sharepoint graph API
        /// </summary>
        /// <param name="webApiUrl">Url of the Web API to call (supposed to return Json)</param>
        /// <param name="accessToken">Access token used as a bearer security token to call the Web API</param>
        /// <param name="existingLibraryId">Library ID of sharepoint document library where file exits</param>
        /// <param name="FileName">File name which is going to copy between libraries</param>
        /// <param name="LibraryName">Destination library name where file is going to copy</param>
        /// <param name="processResult">Callback used to process the result of the call to the Web API</param>
        /// <returns></returns>
        public async Task CopyFile(string webApiUrl, string accessToken, string existingLibraryId, string FileName, string LibraryName, Action<JObject> processResult)
        {
            if (!string.IsNullOrEmpty(accessToken))
            {
                var defaultRequetHeaders = HttpClient.DefaultRequestHeaders;
                if (defaultRequetHeaders.Accept == null || !defaultRequetHeaders.Accept.Any(m => m.MediaType == "application/json"))
                {
                    HttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                }
                defaultRequetHeaders.Authorization = new AuthenticationHeaderValue("bearer", accessToken);

                HttpResponseMessage response = await HttpClient.GetAsync(webApiUrl);
                if (response.IsSuccessStatusCode)
                {
                    string json = await response.Content.ReadAsStringAsync();
                    JObject result = JsonConvert.DeserializeObject(json) as JObject;

                    string libraryId = "";
                    for (var i = 0; i < result["value"].Count(); i++)
                    {
                        if (result["value"][i]["name"].Value<string>() == LibraryName)
                        {
                            libraryId = result["value"][i]["id"].Value<string>();
                        }
                    }

                    if (!(String.IsNullOrEmpty(libraryId)) && libraryId != "")
                    {
                        var parentReference = new ItemReference
                        {
                            Id = libraryId
                        };

                        JObject body = (JObject)JToken.FromObject(parentReference);
                        HttpContent httpContent = new StringContent(JsonConvert.SerializeObject(body), Encoding.UTF8, "application/json");

                        response = await HttpClient.PostAsync("https://graph.microsoft.com/v1.0/drives/" + existingLibraryId + "/root:/" + FileName + ":/copy", httpContent);

                        if (response.IsSuccessStatusCode)
                        {
                            json = await response.Content.ReadAsStringAsync();
                            result = JsonConvert.DeserializeObject(json) as JObject;
                            Console.ForegroundColor = ConsoleColor.Gray;
                            processResult(result);
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine($"Failed to call the Web Api: {response.StatusCode}");
                            string content = await response.Content.ReadAsStringAsync();

                            // Note that if you got reponse.Code == 403 and reponse.content.code == "Authorization_RequestDenied"
                            // this is because the tenant admin as not granted consent for the application to call the Web API
                            Console.WriteLine($"Content: {content}");
                        }
                    }
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Failed to call the Web Api: {response.StatusCode}");
                    string content = await response.Content.ReadAsStringAsync();

                    // Note that if you got reponse.Code == 403 and reponse.content.code == "Authorization_RequestDenied"
                    // this is because the tenant admin as not granted consent for the application to call the Web API
                    Console.WriteLine($"Content: {content}");
                }
                Console.ResetColor();
            }
        }

        /// <summary>
        /// Get all files and folders name using SharePoint graph API
        /// </summary>
        /// <param name="webApiUrl">Url of the Web API to call (supposed to return Json)</param>
        /// <param name="accessToken">Access token used as a bearer security token to call the Web API</param>
        /// <param name="processResult">Callback used to process the result of the call to the Web API</param>
        /// <returns></returns>
        public async Task GetAllFilesandFolders(string webApiUrl, string accessToken, Action<JObject> processResult)
        {
            if (!string.IsNullOrEmpty(accessToken))
            {
                var defaultRequetHeaders = HttpClient.DefaultRequestHeaders;
                if (defaultRequetHeaders.Accept == null || !defaultRequetHeaders.Accept.Any(m => m.MediaType == "application/json"))
                {
                    HttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                }
                defaultRequetHeaders.Authorization = new AuthenticationHeaderValue("bearer", accessToken);

                HttpResponseMessage response = await HttpClient.GetAsync(webApiUrl);
                if (response.IsSuccessStatusCode)
                {
                    string json = await response.Content.ReadAsStringAsync();
                    JObject result = JsonConvert.DeserializeObject(json) as JObject;
                    Console.ForegroundColor = ConsoleColor.Gray;
                    processResult(result);
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Failed to call the Web Api: {response.StatusCode}");
                    string content = await response.Content.ReadAsStringAsync();

                    // Note that if you got reponse.Code == 403 and reponse.content.code == "Authorization_RequestDenied"
                    // this is because the tenant admin as not granted consent for the application to call the Web API
                    Console.WriteLine($"Content: {content}");
                }
                Console.ResetColor();
            }
        }

        /// <summary>
        /// delete files on sharepoint site library, using SharePoint graph API
        /// </summary>
        /// <param name="webApiUrl">Url of the Web API to call (supposed to return Json)</param>
        /// <param name="accessToken">Access token used as a bearer security token to call the Web API</param>
        /// <param name="processResult">Callback used to process the result of the call to the Web API</param>
        /// <returns></returns>
        public async Task DeleteFile(string webApiUrl, string accessToken, Action<JObject> processResult)
        {
            if (!string.IsNullOrEmpty(accessToken))
            {
                var defaultRequetHeaders = HttpClient.DefaultRequestHeaders;
                if (defaultRequetHeaders.Accept == null || !defaultRequetHeaders.Accept.Any(m => m.MediaType == "application/json"))
                {
                    HttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                }
                defaultRequetHeaders.Authorization = new AuthenticationHeaderValue("bearer", accessToken);

                HttpResponseMessage response = await HttpClient.DeleteAsync(webApiUrl);
                if (response.IsSuccessStatusCode)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("File deleted!!!");
                    Console.ResetColor();

                    string json = await response.Content.ReadAsStringAsync();
                    JObject result = JsonConvert.DeserializeObject(json) as JObject;
                    Console.ForegroundColor = ConsoleColor.Gray;
                    processResult(result);
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Failed to call the Web Api: {response.StatusCode}");
                    string content = await response.Content.ReadAsStringAsync();

                    // Note that if you got reponse.Code == 403 and reponse.content.code == "Authorization_RequestDenied"
                    // this is because the tenant admin as not granted consent for the application to call the Web API
                    Console.WriteLine($"Content: {content}");
                }
                Console.ResetColor();
            }
        }

        /// <summary>
        /// create folder on sharepoint site library, using sharepoint graph api
        /// </summary>
        /// <param name="webApiUrl">Url of the Web API to call (supposed to return Json)</param>
        /// <param name="accessToken">Access token used as a bearer security token to call the Web API</param>
        /// <param name="body">JSON structure of sending data, which is required in post request</param>
        /// <param name="processResult">Callback used to process the result of the call to the Web API</param>
        /// <returns></returns>
        public async Task CreateFolder(string webApiUrl, string accessToken, JObject body, Action<JObject> processResult)
        {
            if (!string.IsNullOrEmpty(accessToken))
            {
                var defaultRequetHeaders = HttpClient.DefaultRequestHeaders;
                if (defaultRequetHeaders.Accept == null || !defaultRequetHeaders.Accept.Any(m => m.MediaType == "application/json"))
                {
                    HttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                }
                defaultRequetHeaders.Authorization = new AuthenticationHeaderValue("bearer", accessToken);

                HttpContent httpContent = new StringContent(JsonConvert.SerializeObject(body), Encoding.UTF8, "application/json");

                HttpResponseMessage response = await HttpClient.PostAsync(webApiUrl, httpContent);
                if (response.IsSuccessStatusCode)
                {
                    string json = await response.Content.ReadAsStringAsync();
                    JObject result = JsonConvert.DeserializeObject(json) as JObject;
                    Console.ForegroundColor = ConsoleColor.Gray;
                    processResult(result);
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Failed to call the Web Api: {response.StatusCode}");
                    string content = await response.Content.ReadAsStringAsync();

                    // Note that if you got reponse.Code == 403 and reponse.content.code == "Authorization_RequestDenied"
                    // this is because the tenant admin as not granted consent for the application to call the Web API
                    Console.WriteLine($"Content: {content}");
                }
                Console.ResetColor();
            }
        }
    }
}
