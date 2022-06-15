using MailKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LaRottaO.CSharp.Utils.Email
{
    public class EmailMessage
    {
        public UniqueId uniqueId { get; set; }
        public String from { get; set; }
        public String subject { get; set; }
        public DateTimeOffset receivedDate { get; set; }
        public String body { get; set; }
        public List<String> downloadedAttachments { get; set; }

        public EmailMessage()
        {
            this.downloadedAttachments = new List<String>();
        }
    }
}