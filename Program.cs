using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;

namespace ReadGMailMails
{
    class Program
    {
        static void Main(string[] args)
        {
            GmailApplicationService GmailApplicationService = new GmailApplicationService();

            //string outputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "Teste");
            string outputPath = ConfigurationManager.AppSettings["outputPath_files"].ToString();

            if(!Directory.Exists(outputPath))
            {
                Directory.CreateDirectory(outputPath);
            }

            #region [ Testes de Mensagens ]
            var ListMessages = GmailApplicationService.ListAllMessages(false, new List<string> { "CATEGORY_PERSONAL" }, 100);

            foreach (var messageId in ListMessages)
            {
                // carrega a mensagem
                Google.Apis.Gmail.v1.Data.Message message = GmailApplicationService.GetMessage(messageId.Id);

                // baixa as imagens
                //GmailApplicationService.GetImagesAttachments(message, outputPath);
            }



            #endregion

            #region [ Teste de Labels ]
            //var Labels = GmailApplicationService.GetLabels();

            //foreach (var item in Labels.Where(c => c.Type == "system").OrderBy(c => c.Type == "system").ThenBy(c => c.Name))
            //{
            //    Console.WriteLine($"► {item.Name} - [ID_LABEL]: {item.Id}");

            //    if (item.Id == "CATEGORY_PERSONAL")
            //    {
            //        var label = GmailApplicationService.GetLabel(item.Id);

            //        string json = Newtonsoft.Json.JsonConvert.SerializeObject(label, Newtonsoft.Json.Formatting.Indented);

            //        Console.WriteLine(json);
            //    }
            //}
            #endregion

            #region [ Teste de Profile ]
            //var profile = GmailApplicationService.GetUserProfile();
            #endregion



            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        
    }
}
