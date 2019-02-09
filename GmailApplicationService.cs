using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;

namespace ReadGMailMails
{
    public class GmailApplicationService
    {
        public GmailService GmailService { get; set; }

        public GmailApplicationService()
        {
            this.GmailService = new GmailAuthorization().Authorization();
        }

        private int GetMessageCount(int? countMessages)
        {
            int count = 0;

            if (countMessages.HasValue) // tem valor
            {
                if (countMessages > 0) // é maior que zero
                {
                    count = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(countMessages.Value) / Convert.ToDouble(100)) * 100); // converte para o multiplo de 100 mais próximo ou exato
                }
            }

            return count;
        }
        private byte[] FromBase64ForUrlString(string base64ForUrlInput)
        {
            int padChars = (base64ForUrlInput.Length % 4) == 0 ? 0 : (4 - (base64ForUrlInput.Length % 4));
            StringBuilder result = new StringBuilder(base64ForUrlInput, base64ForUrlInput.Length + padChars);
            result.Append(String.Empty.PadRight(padChars, '='));
            result.Replace('-', '+');
            result.Replace('_', '/');
            return Convert.FromBase64String(result.ToString());
        }
        private string Base64UrlDecode(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return "<strong>Message body was not returned from Google</strong>";
            }

            string InputStr = input.Replace("-", "+").Replace("_", "/");

            return System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(InputStr));
        }



        /// <summary>
        /// Gets the current user's Gmail profile.
        /// </summary>
        /// <param name="userId">The user's email address. The special value "me" can be used to indicate the authenticated user.</param>
        public Profile GetUserProfile()
        {
            Profile profile = new Profile();

            try
            {
                profile = GmailService.Users.GetProfile("me").Execute();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            return profile;
        }


        /// <summary>
        /// Lists ALL messages in the user's mailbox according to the parameters
        /// </summary>
        /// <param name="userId">The user's email address. The special value "me" can be used to indicate the authenticated user.</param>
        /// <param name="includeSpamTrash">Include messages from SPAM and TRASH in the results. (Default: false)</param>
        /// <param name="labelIds">Only return messages with labels that match all of the specified label IDs.</param>
        /// <param name="maxResults">Maximum number of messages to return.</param>
        /// <param name="query">Only return messages matching the specified query. Supports the same query format as the Gmail search box. For example, "from:someuser@example.com rfc822msgid:<somemsgid@example.com> is:unread". Parameter cannot be used when accessing the api using the gmail.metadata scope.</param>
        /// <returns></returns>
        public List<Message> ListAllMessages(bool? includeSpamTrash = null, IEnumerable<string> labelIds = null, int? maxResults = null, string query = null)
        {
            List<Message> messages = new List<Message>();

            UsersResource.MessagesResource.ListRequest request = GmailService.Users.Messages.List("me");
            request.Q = query == null ? null : query;
            request.IncludeSpamTrash = includeSpamTrash == null ? false : includeSpamTrash;
            request.LabelIds = labelIds == null ? null : new Google.Apis.Util.Repeatable<string>(labelIds);
            request.MaxResults = maxResults == null ? null : maxResults;

            do
            {
                try
                {
                    ListMessagesResponse response = request.Execute();

                    messages.AddRange(response.Messages);

                    request.PageToken = response.NextPageToken;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                }
            } while (!String.IsNullOrEmpty(request.PageToken));

            return messages;
        }


        /// <summary>
        /// Object that contains: List of messages, Token to retrieve the next page of results in the list and Estimated total number of results.
        /// </summary>        
        /// <param name="userId">The user's email address. The special value "me" can be used to indicate the authenticated user.</param>
        /// <param name="includeSpamTrash">Include messages from SPAM and TRASH in the results. (Default: false)</param>
        /// <param name="labelIds">Only return messages with labels that match all of the specified label IDs.</param>
        /// <param name="maxResults">Maximum number of messages to return.</param>
        /// <param name="pageToken">Page token to retrieve a specific page of results in the list.</param>
        /// <param name="query">Only return messages matching the specified query. Supports the same query format as the Gmail search box. For example, "from:someuser@example.com rfc822msgid:<somemsgid@example.com> is:unread". Parameter cannot be used when accessing the api using the gmail.metadata scope.</param>
        /// <returns></returns>
        public ListMessagesResponse ListMessages(bool? includeSpamTrash = null, IEnumerable<string> labelIds = null, int? maxResults = null, string pageToken = null, String query = null)
        {
            ListMessagesResponse response = null;

            UsersResource.MessagesResource.ListRequest request = GmailService.Users.Messages.List("me");
            request.Q = query == null ? null : query;
            request.IncludeSpamTrash = includeSpamTrash == null ? false : includeSpamTrash;
            request.LabelIds = labelIds == null ? null : new Google.Apis.Util.Repeatable<string>(labelIds);
            request.MaxResults = maxResults == null ? null : maxResults;
            request.PageToken = pageToken == null ? null : pageToken;

            try
            {
                response = request.Execute();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            return response;
        }


        /// <summary>
        /// Retrieve a Message by ID.
        /// </summary>        
        /// <param name="userId">User's email address. The special value "me" can be used to indicate the authenticated user.</param>
        /// <param name="messageId">ID of Message to retrieve.</param>
        public Message GetMessage(string messageId)
        {
            try
            {
                return GmailService.Users.Messages.Get("me", messageId).Execute();
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
            }

            return null;
        }


        /// <summary>
        /// Get and store images files attachment from Message
        /// </summary>        
        /// <param name="userId">User's email address. The special value "me" can be used to indicate the authenticated user.</param>
        /// <param name="messageId">ID of Message containing attachment.</param>
        /// <param name="outputDir">Directory used to store attachments.</param>
        public void GetImagesAttachments(Message message, string outputDir)
        {
            try
            {
                if (message.Payload.Parts != null)
                {
                    IList<MessagePart> parts = message.Payload.Parts;

                    foreach (MessagePart part in parts)
                    {
                        if (!String.IsNullOrEmpty(part.Filename) && part.MimeType.Contains("image/"))
                        {
                            String attId = part.Body.AttachmentId;

                            if (attId != null)
                            {
                                MessagePartBody attachPart = GmailService.Users.Messages.Attachments.Get("me", message.Id, attId).Execute();

                                // Converting from RFC 4648 base64 to base64url encoding
                                // see http://en.wikipedia.org/wiki/Base64#Implementations_and_history
                                String attachData = attachPart.Data.Replace('-', '+');
                                attachData = attachData.Replace('_', '/');

                                byte[] data = Convert.FromBase64String(attachData);
                                File.WriteAllBytes(Path.Combine(outputDir, part.Filename), data);
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"An error occurred: {e.Message} - [{message.Id}]");
            }
        }


        /// <summary>
        /// Get and store pdf files attachment from Message
        /// </summary>        
        /// <param name="userId">User's email address. The special value "me" can be used to indicate the authenticated user.</param>
        /// <param name="messageId">ID of Message containing attachment.</param>
        /// <param name="outputDir">Directory used to store attachments.</param>
        public void GetPdfAttachments(Message message, string outputDir)
        {
            try
            {
                if (message.Payload.Parts != null)
                {
                    IList<MessagePart> parts = message.Payload.Parts;

                    foreach (MessagePart part in parts)
                    {
                        if (!String.IsNullOrEmpty(part.Filename) && part.MimeType == "application/pdf")
                        {
                            String attId = part.Body.AttachmentId;

                            if (attId != null)
                            {
                                MessagePartBody attachPart = GmailService.Users.Messages.Attachments.Get("me", message.Id, attId).Execute();

                                // Converting from RFC 4648 base64 to base64url encoding
                                // see http://en.wikipedia.org/wiki/Base64#Implementations_and_history
                                String attachData = attachPart.Data.Replace('-', '+');
                                attachData = attachData.Replace('_', '/');

                                byte[] data = Convert.FromBase64String(attachData);
                                File.WriteAllBytes(Path.Combine(outputDir, part.Filename), data);
                            }
                        }
                    }
                }


            }
            catch (Exception e)
            {
                Console.WriteLine($"An error occurred: {e.Message} - [{message.Id}]");
            }
        }


        /// <summary>
        /// Get and store html content from Message
        /// </summary>        
        /// <param name="userId">User's email address. The special value "me" can be used to indicate the authenticated user.</param>
        /// <param name="messageId">ID of Message containing attachment.</param>
        /// <param name="outputDir">Directory used to store attachments.</param>
        public void GetMessageAsHtml(Message message, string outputDir)
        {
            try
            {
                if (message.Payload.Parts != null)
                {
                    IList<MessagePart> parts = message.Payload.Parts;

                    foreach (MessagePart part in parts)
                    {
                        if (part.MimeType == "text/html")
                        {
                            string filename = $"pagina_{DateTime.Now.ToString("ddMMyyyy_hhmmss")}.html";
                            byte[] data = FromBase64ForUrlString(part.Body.Data);
                            string decodedString = Encoding.UTF8.GetString(data);
                            File.WriteAllBytes(Path.Combine(outputDir, filename), data);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"An error occurred: {e.Message} - [{message.Id}]");
            }
        }


        /// <summary>
        /// List the labels in the user's mailbox.
        /// <param name="userId">User's email address. The special value "me" can be used to indicate the authenticated user.</param>
        /// </summary>
        public IList<Label> GetLabels()
        {
            IList<Label> Labels = new List<Label>();

            try
            {
                ListLabelsResponse response = GmailService.Users.Labels.List("me").Execute();

                Labels = response.Labels;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            return Labels;
        }


        /// <summary>
        /// Get specified Label.
        /// </summary>
        /// <param name="userId">User's email address. The special value "me" can be used to indicate the authenticated user.</param>
        /// <param name="labelId">ID of Label to get.</param>
        /// <returns></returns>
        public Label GetLabel(string labelId)
        {
            Label Label = new Label();

            try
            {
                Label = GmailService.Users.Labels.Get("me", labelId).Execute();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            return Label;
        }


    }
}
