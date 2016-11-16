using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using System;
using System.Configuration;
using System.Security.Cryptography.X509Certificates;

namespace GoogleSpreadSheetReader.Provider
{
    public class ServiceCredentialsProvider
    {
        public static ServiceAccountCredential GetServiceCredentials()
        {
            ServiceAccountCredential credential;

            string[] scopes = { SheetsService.Scope.SpreadsheetsReadonly };
            string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            var keyFilePath = ConfigurationManager.AppSettings["P12Certificate"]; 
 
             var serviceAccountEmail = ConfigurationManager.AppSettings["ServiceAccountEmail"];

            var certificate = new X509Certificate2(keyFilePath, ConfigurationManager.AppSettings["P12CertificatePassword"], X509KeyStorageFlags.Exportable);
            credential = new ServiceAccountCredential(new ServiceAccountCredential.Initializer(serviceAccountEmail)
            {
                Scopes = scopes
            }.FromCertificate(certificate));

            return credential;
        }
    }
}
