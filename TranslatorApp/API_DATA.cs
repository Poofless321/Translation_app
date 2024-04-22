using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.IO;
using TranslatorApp;

namespace TranslatorApp
{
    internal class API_DATA
    {
        public readonly HttpClient client = new HttpClient();
        public string? API_KEY = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
        public string url;

        public API_DATA(string url, TimeSpan? requestTimeout = null)
        {
            this.url = url;

            // Set a custom timeout if specified, otherwise use the default
            if (requestTimeout != null)
            {
                client.Timeout = (TimeSpan)requestTimeout;
            }
        }


        public void headers()
        {
            if (!client.DefaultRequestHeaders.Contains("Authorization"))
            {
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {API_KEY}"); // Add header only if it doesn't already exist
            }
        }

        public dynamic data(string data)
        {
            return new
            {
                model = "gpt-3.5-turbo-0125",
                messages = new[]
{
                 new { role = "system", content = "You are a helpful assistant designed to translate text and output JSON and maintain  the exact same format." },
                new { role = "user", content = $"Translate the following game bubbles to english in JSON there should be no nesting) and maintain the exact same text format only translate japanese words:\n{data}" }
            }, // $"Translate the following text bubbles to english in JSON (use the following format EX: {{'1': translation, '2': translation}} there should be no nesting and if applicable keep special character codes in their orginal positions): \n{data}" }
                temperature = 0.4
            };
        }

        public async Task<String?> API(string text)
        {
            //Make the API call
            string json = JsonConvert.SerializeObject(data(text));
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            //wait and verify reponse
            var response = await client.PostAsync(url, content);
            response.EnsureSuccessStatusCode();

            //Get the response
            string responseBody = await response.Content.ReadAsStringAsync();
            dynamic openAIResponse = JsonConvert.DeserializeObject<dynamic>(responseBody);

            if (openAIResponse == null)
            {
                return "Response was Null";
            }

            return openAIResponse.choices[0].message.content;
        }


    }
}
