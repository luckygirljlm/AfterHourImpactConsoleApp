using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AfterHourConsoleApplication.SDK
{
    public class Message
    {
        public string InternetMsgId { get; set; }
        public string ItemId { get; set; }
        public string ConversationId { get; set; }
        public DateTime SentDateTime { get; set; }
        public string Subject { get; set; }
        public string InReplyTo { get; set; }
        public Person Sender { get; set; }
    }

    public class ReplyMessage : Message
    {
        public string InReplyToItemId { get; set; }
        public DateTime InReplyToSentDateTime { get; set; }

    }

    public class MessageStat
    {

    }

    public class Person
    {
        public string Name { get; set; }
        public string Title { get; set; }
        public string SmtpAddress { get; set; }
    }

    public class Conversation
    {
        public string ConversationId { get; set; }
        public string ConversationTopic { get; set; }
        public string[] ItemsSent { get; set; }
        public string[] ItemsAll { get; set; }

    }

    public class WorkingHours
    {
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public long TimeZoneBias { get; set; }  //China -480
        public int[] WorkDays { get; set; } //default [0, 1, 1, 1, 1, 1, 0], start from Sunday

    }
}
