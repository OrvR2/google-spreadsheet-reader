using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using GoogleSpreadSheetReader.Model;
using GoogleSpreadSheetReader.Provider;
using Newtonsoft.Json;
using System.Configuration;
using System.Xml;

namespace MySpreadsheetIntegration
{
    class Program
    {
        static void Main(string[] args)
        {
            var credential = ServiceCredentialsProvider.GetServiceCredentials();

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ConfigurationManager.AppSettings["ApplicationName"],
            });

            var spreadsheetId = ConfigurationManager.AppSettings["SpreadSheetId"];
            var range = ConfigurationManager.AppSettings["Range"];
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

            var response = request.Execute();
            var values = response.Values;
            var json = JsonConvert.SerializeObject(new Columns() { Rows = values });

            XmlDocument doc = JsonConvert.DeserializeXmlNode(json, "Columns");
            doc.Save(ConfigurationManager.AppSettings["OutputXmlFile"]);            
        }
    }
}