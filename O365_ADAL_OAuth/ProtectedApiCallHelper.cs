/*
 * (ONLY FOR CALENDAR EVENTS)
 * 
 * This solution comes from a daemon console application. 
 * In that, this class was used for calling to a single API, get(CallWebApiAndProcessResultASync).
 * 
 * As for this project requirement, 
 * the developer did make a change to this class to perform crud operation on Calendar Events.
 * 
 * Now this class perform post, patch(put), delete and get(CallWebApiAndProcessResultASync) on Calendar event via graph API.
 * In return, it will provide a JSON response.
 */
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace O365_ADAL_OAuth
{
    /// <summary>
    /// Helper class to call a protected API and process its result
    /// </summary>
    public class ProtectedApiCallHelper
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="httpClient">HttpClient used to call the protected API</param>
        public ProtectedApiCallHelper(HttpClient httpClient)
        {
            HttpClient = httpClient;
        }

        protected HttpClient HttpClient { get; private set; }

        /// <summary>
        /// Calls the protected Web API and processes the result
        /// </summary>
        /// <param name="webApiUrl">Url of the Web API to call (supposed to return Json)</param>
        /// <param name="accessToken">Access token used as a bearer security token to call the Web API</param>
        /// <param name="processResult">Callback used to process the result of the call to the Web API</param>
        public async Task CallWebApiAndProcessResultASync(string webApiUrl, string accessToken, Action<JObject> processResult)
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
        /// post method and processes the result
        /// </summary>
        /// <param name="webApiUrl">Url of the Web API to call (supposed to return Json)</param>
        /// <param name="accessToken">Access token used as a bearer security token to call the Web API</param>
        /// <param name="processResult">Callback used to process the result of the call to the Web API</param>
        /// <param name="body">JSON structure of sending data, which is required in post request</param>
        /// <returns></returns>
        public async Task post(string webApiUrl, string accessToken, JObject body, Action<JObject> processResult)
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

        /// <summary>
        /// delete method and processes the result
        /// </summary>
        /// <param name="webApiUrl">Url of the Web API to call (supposed to return Json)</param>
        /// <param name="accessToken">Access token used as a bearer security token to call the Web API</param>
        /// <param name="processResult">Callback used to process the result of the call to the Web API</param>
        /// <returns></returns>
        public async Task delete(string webApiUrl, string accessToken, Action<JObject> processResult)
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

                    //delete method to delete event using specific subject
                    response = await HttpClient.DeleteAsync(webApiUrl.Split('?')[0] + "/" + result["value"][0]["id"].Value<string>());

                    if (response.IsSuccessStatusCode)
                    {

                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Event deleted!!!");
                        Console.ResetColor();
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
        /// update method and processes the result
        /// </summary>
        /// <param name="webApiUrl">Url of the Web API to call (supposed to return Json)</param>
        /// <param name="accessToken">Access token used as a bearer security token to call the Web API</param>
        /// <param name="processResult">Callback used to process the result of the call to the Web API</param>
        /// <param name="body">JSON structure of sending data, which is required in put request</param>
        /// <returns></returns>
        public async Task put(string webApiUrl, string accessToken, JObject body, Action<JObject> processResult)
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

                    //update method to update event using specific subject
                    HttpContent httpContent = new StringContent(JsonConvert.SerializeObject(body), Encoding.UTF8, "application/json");
                    response = await HttpClient.PatchAsync(webApiUrl.Split('?')[0] + "/" + result["value"][0]["id"].Value<string>(), httpContent);
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
