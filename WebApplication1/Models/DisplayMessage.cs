using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebApplication1.Models
{
    public class DisplayMessage
    {
        public string Subject { get; set; }
        public DateTimeOffset ReceivedDateTime { get; set; }
        public string From { get; set; }
        public string Body { get; set; }
        public string Recipient { get; set; }
        public string EmailStatus { get; set; }

        public DisplayMessage(string subject, DateTimeOffset? dateTimeReceived,
            Microsoft.Office365.OutlookServices.Recipient recipient, Microsoft.Office365.OutlookServices.ItemBody body, Microsoft.Office365.OutlookServices.EmailAddress from)
        {
            this.Subject = subject;
            this.ReceivedDateTime = (DateTimeOffset)dateTimeReceived;
            //this.From = from != null ? string.Format("{0} ({1})", from.EmailAddress.Name, from.EmailAddress.Address) | "EMPTY";
            this.From = from.Address;
            this.Body = body.Content;
            this.Recipient = recipient.EmailAddress.Address;
        }
    }
}