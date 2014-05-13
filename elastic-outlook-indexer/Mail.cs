using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace elastic_outlook_indexer
{
    public class Mail
    {
        public string BCC { get; set; }
        public string Body { get; set; }
        public string BodyFormat { get; set; }
        public string Categories { get; set; }
        public string CC { get; set; }
        public DateTime CreationTime { get; set; }
        public string EntryId { get; set; }
        public string FlagIcon { get; set; }
        public DateTime FlagDueBy { get; set; }
        public string FlagRequest { get; set; }
        public string FlagStatus { get; set; }
        public string HtmlBody { get; set; }
        public DateTime LastModificationTime { get; set; }
        public string ReceivedByEntryId { get; set; }
        public string ReceivedByName { get; set; }
        public DateTime ReceivedTime { get; set; }
        public int InternetCodepage { get; set; }
        public string ConversationID { get; set; }
        public string ConversationIndex { get; set; }
        public string ConversationTopic { get; set; }
        public IList<Recipient> Recipients { get; set; }
        public string ReplyRecipientNames { get; set; }
        public IList<Recipient> ReplyRecipients { get; set; }
        public byte[] RTFBody { get; set; }
        public bool Saved { get; set; }
        public AddressEntry Sender { get; set; }
        public string SenderEmailAddress { get; set; }
        public string SenderEmailType { get; set; }
        public string SenderName { get; set; }
        public bool Sent { get; set; }
        public DateTime SentOn { get; set; }
        public int Size { get; set; }
        public string Subject { get; set; }
        public string To { get; set; }
        public string VotingOptions { get; set; }
        public string VotingResponse { get; set; }
    }
}
