using MicrosoftGraphEmails.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph.Models;


namespace MicrosoftGraphEmails
{
    class Program
    {
        static void Main(string[] args)
        {
            //Secrets contain creds for Graph
            var config = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddUserSecrets<Program>()
                .Build();

            MicrosoftGraphEmailService msGraphService = new MicrosoftGraphEmailService(config);


            msGraphService.SendMessageAsync("mailbox@domain.com",
                new List<string>() { "email@domain.com" },
                new List<string>() { "email@domain.com" },
                new List<string>() { "email@domain.com" },
                "Subject Text",
                "Body Text",
                Importance.High,
                new List<string>() { "/path/to/file.ext" }
                ).GetAwaiter().GetResult();


        }
    }
}