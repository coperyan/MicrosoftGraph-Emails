using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MicrosoftGraphEmails.Models
{
    public class Email
    {
        public string mailbox { get; set; }
        public List<string> to_recipients { get; set; }

        public List<string> cc_recipients { get; set; }

        public List<string> bcc_recipients { get; set; }

        public string subject { get; set; }

        public string body { get; set; }

        public List<string> file_attachments { get; set; }

        public string importance { get; set; }
    }
}
