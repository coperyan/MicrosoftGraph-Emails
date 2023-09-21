using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using MicrosoftGraphEmails.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MicrosoftGraphEmails.Services
{
    public class MicrosoftGraphEmailService
    {
        private readonly IConfiguration _config;

        public MicrosoftGraphEmailService(IConfiguration config)
        {
            _config = config;
        }

        public async Task SendAsync(Email email)
        {

        }
    }

}
