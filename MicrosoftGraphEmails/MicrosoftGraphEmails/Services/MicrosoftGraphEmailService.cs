using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;


namespace MicrosoftGraphEmails.Services
{
    public class MicrosoftGraphEmailService
    {
        private readonly IConfiguration _config;
        private readonly GraphServiceClient _graphClient;

        public MicrosoftGraphEmailService(IConfiguration config)
        {
            _config = config;
            _graphClient = new GraphServiceClient(
                new ClientSecretCredential(
                    _config["tenantId"],
                    _config["clientId"],
                    _config["clientSecret"]
                    )
                );
        }

        private List<Recipient> genEmailList(List<string> emails)
        {
            List<Recipient> emailList = new List<Recipient>();
            foreach(var recipient in emails)
            {
                emailList.Add(new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = recipient,
                    },
                });
            }
            return emailList;
        }

        private async Task<string> CreateMessage(string mailbox, List<string> to_recipients, List<string> cc_recipients, List<string> bcc_recipients, string subject, string body, Importance importance = Importance.Normal)
        {
            var requestBody = new Message
            {
                Subject = subject,
                Importance = importance,
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = body
                },
                ToRecipients = genEmailList(to_recipients),
                CcRecipients = genEmailList(cc_recipients),
                BccRecipients = genEmailList(bcc_recipients),
            };

            var result = await _graphClient.Users[mailbox].Messages.PostAsync(requestBody);
            Console.WriteLine(result.Id);
            return result.Id;

        }

        public async Task UploadLargeAttachment(string mailbox, string messageId, string attachment)
        {
            using var fileStream = File.OpenRead(attachment);
            var largeAttachment = new AttachmentItem
            {
                AttachmentType = AttachmentType.File,
                Name = Path.GetFileName(attachment),
                Size = fileStream.Length,
            };

            var uploadSessionRequestBody = new Microsoft.Graph.Users.Item.Messages.Item.Attachments.CreateUploadSession.CreateUploadSessionPostRequestBody
            {
                AttachmentItem = largeAttachment
            };

            var uploadSession = await _graphClient
                .Users[mailbox]
                .Messages[messageId]
                .Attachments
                .CreateUploadSession
                .PostAsync(uploadSessionRequestBody);

            int maxSliceSize = 320 * 1024;
            var fileUploadTask =
                new LargeFileUploadTask<FileAttachment>(uploadSession, fileStream, maxSliceSize);

            var totalLength = fileStream.Length;

            IProgress<long> progress = new Progress<long>(prog =>
            {
                Console.WriteLine(string.Format("Uploaded {0} bytes of {1}..", prog, totalLength));
            });

            try
            {
                var uploadResult = await fileUploadTask.UploadAsync(progress);
                Console.WriteLine(uploadResult.UploadSucceeded ? "Upload complete" : "Upload failed");
            }
            catch (ODataError ex)
            {
                Console.WriteLine($"Error uploading: {ex.Error?.Message}");
            }


        }

        public async Task SendMessageAsync(string mailbox, List<string> to_recipients, List<string> cc_recipients, List<string> bcc_recipients, string subject, string body, Importance importance = Importance.Normal, List<string> file_attachments = null)
        {
            var messageId = await CreateMessage(mailbox, to_recipients, cc_recipients, bcc_recipients, subject, body, importance);
            Console.WriteLine(string.Format("Got message ID:{0}", messageId));

            foreach(var attachment in file_attachments)
            {
                await UploadLargeAttachment(mailbox, messageId, attachment);
                Console.WriteLine(string.Format("Uploaded attachment {0}", attachment));
            }

            //Needs dev to handle HTML body + image (content) attachments

            await _graphClient.Users[mailbox].Messages[messageId].Send.PostAsync();
            Console.WriteLine("Sent message..");

        }
    }

}
