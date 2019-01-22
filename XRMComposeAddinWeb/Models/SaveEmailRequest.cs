using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace XRMComposeAddinWeb.Models
{
    public class SaveEmailRequest
    {
        public string Subject { get; set; }
        public string Message { get; set; }
        public string Sender { get; set; }
        public string To { get; set; }
        public string InOut { get; set; }
        public string Category { get; set; }
        public string CaseID { get; set; }
        public string RelatedItemListId { get; set; }
        public string CreatedDateTime { get; set; }
    }
}