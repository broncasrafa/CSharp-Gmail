using System.Configuration;
using System.IO;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace ReadGMailMails
{
    public class GmailAuthorization
    {
        public GmailService Authorization()
        {
            string ApplicationName = "MyMails";

            string[] Scopes = {
                GmailService.Scope.GmailReadonly,
                GmailService.Scope.GmailLabels,
                GmailService.Scope.GmailMetadata,
                GmailService.Scope.GmailCompose,
                GmailService.Scope.GmailModify
            };

            UserCredential credential;

            var filePath = Directory.GetFiles(ConfigurationManager.AppSettings["credentials_path"].ToString())[0];

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets,
                                                                         Scopes,
                                                                         "user",
                                                                         CancellationToken.None,
                                                                         new FileDataStore(credPath, true)).Result;
            }

            // Create Gmail API service.
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }
    }
}
