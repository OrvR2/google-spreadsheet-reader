using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using GoogleSpreadSheetReader.Model;
using GoogleSpreadSheetReader.Provider;
using Newtonsoft.Json;
using System;
using System.Configuration;
using System.Xml;

namespace MySpreadsheetIntegration
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Title = "Google Reader";

            try
            {
                Console.WriteLine("1) Getting Service Credentials...");
                var credential = ServiceCredentialsProvider.GetServiceCredentials();
                Console.WriteLine("A Service Credential was successfully created !");

                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ConfigurationManager.AppSettings["ApplicationName"],
                });

                var spreadsheetId = ConfigurationManager.AppSettings["SpreadSheetId"];
                var range = ConfigurationManager.AppSettings["Range"];
                SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                Console.WriteLine("2) Getting SpreadSheet from Google Drive...");
                var response = request.Execute();
                var values = response.Values;
                Console.WriteLine("SpreadSheet values are received sucessfully!");

                Console.WriteLine("3) Serializing information received from SpreadSheet....");
                var json = JsonConvert.SerializeObject(new Columns() { Rows = values });
                XmlDocument doc = JsonConvert.DeserializeXmlNode(json, "Columns");
                doc.Save(ConfigurationManager.AppSettings["OutputXmlFile"]);
                Console.WriteLine("4) We finished!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Error Ocurred: {0}", ex.Message));
            }
        }
    }
}