using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AfterHourConsoleApplication.SDK;
using System.Xml;

namespace AfterHourConsoleApplication
{
    class ParseUtils
    {
        public static Message[] sentMailPropsParsed(string response)
        {
            Message[] result = { };
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(response);

            XmlNodeList messages = xmlDoc.GetElementsByTagName("t:Message");

            for (int i = 0; i < messages.Count; i++)
            {
                string sentTimeTicks = getXMLExactlyOneElementOrThrow(messages.Item(i), "t:DateTimeSent").InnerText;
                DateTime itemSentTime = new DateTime();
                try
                {
                    itemSentTime = Convert.ToDateTime(sentTimeTicks.Substring(0, 10) + " " + sentTimeTicks.Substring(11, 8));
                }
                catch (Exception ex)
                {
                    //Console.WriteLine("Error: DateTimeSent is invalid when parsing response for FindSentItems...Skip this item...");
                    continue;
                }

                string subject = getXMLElementTextContentOrEmpty(messages.Item(i), "t:Subject");

                string inReplyTo = getXMLElementTextContentOrEmpty(messages.Item(i), "t:InReplyTo");

                XmlElement conversationElement = (XmlElement)(getXMLExactlyOneElementOrThrow(messages.Item(i), "t:ConversationId"));
                string conversationId = conversationElement.GetAttribute("Id");

                string internetMsgId = getXMLExactlyOneElementOrThrow(messages.Item(i), "t:InternetMessageId").InnerText;

                XmlElement itemElement = (XmlElement)(getXMLExactlyOneElementOrThrow(messages.Item(i), "t:ItemId"));
                string itemId = itemElement.GetAttribute("Id");

                XmlNode mailboxNode = getXMLExactlyOneElementOrThrow(messages.Item(i), "t:Mailbox");
                string senderName = getXMLElementTextContentOrEmpty(mailboxNode, "t:Name");
                string senderAddress = getXMLElementTextContentOrEmpty(mailboxNode, "t:EmailAddress");
                if (senderName == "" || senderAddress == "")
                {
                    //Console.WriteLine("Error: senderName or senderAddress is empty when parsing response for FindSentItems...Skip this item...");
                    continue;
                }

                if (senderName != "" && Utils.senderName=="") {
                    Utils.senderName = senderName;
                    if (Utils.receipientsConfig.ContainsKey(senderName.Replace(" ", "").ToLower()))
                    {
                        Utils.userWorkingHourConfig = Utils.receipientsConfig[senderName.Replace(" ", "").ToLower()];
                    }
                }

                Message sentMail = new Message
                {
                    InternetMsgId = internetMsgId,
                    ItemId = itemId,
                    ConversationId = conversationId,
                    SentDateTime = itemSentTime,
                    Subject = subject,
                    InReplyTo = inReplyTo,
                    Sender = new Person { Name = senderName, SmtpAddress = senderAddress, Title = "" }
                };

                result = result.Concat(new Message[] { sentMail }).ToArray();
            }

            return result;
        }

        public static Conversation[] joinedConversationPropsParsed(string response)
        {
            Conversation[] result = { };
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(response);

            XmlNodeList conversations = xmlDoc.GetElementsByTagName("Conversation");
            for (int i = 0; i < conversations.Count; i++)
            {
                XmlNode conversationNode = conversations.Item(i);
                XmlElement conversationElement = (XmlElement)(getXMLExactlyOneElementOrThrow(conversationNode, "ConversationId"));
                string conversationId = conversationElement.GetAttribute("Id");

                string conversationTopic = getXMLElementTextContentOrEmpty(conversationNode, "ConversationTopic");

                XmlNode itemsSentNode = getXMLExactlyOneElementOrThrow(conversationNode, "ItemIds");
                string[] itemsSent = parseItems(itemsSentNode);

                XmlNode itemsAllNode = getXMLExactlyOneElementOrThrow(conversationNode, "GlobalItemIds");
                string[] itemsAll = parseItems(itemsAllNode);

                Conversation conv = new Conversation
                {
                    ConversationId = conversationId,
                    ConversationTopic = conversationTopic,
                    ItemsSent = itemsSent,
                    ItemsAll = itemsAll
                };

                result = result.Concat(new Conversation[] { conv }).ToArray();
            }

            return result;
        }

        public static WorkingHours userConfigurationPropsParsed(string response)
        {
            WorkingHours result = null;
            string workDays = "Monday To Friday";
            string defaultStart = "09:00:00";
            string defaultEnd = "18:00:00";
            int[] defaultWorkdays = new int[] { 0, 1, 1, 1, 1, 1, 0 };
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(response);

            XmlNodeList xmlData = xmlDoc.GetElementsByTagName("t:XmlData");
            if (xmlData == null || xmlData.Count == 0)
            {
                result = new WorkingHours
                {
                    StartDate = defaultStart,
                    EndDate = defaultEnd,
                    TimeZoneBias = 0,
                    WorkDays = defaultWorkdays
                };
            }
            else {
                string workHoursBase64 = xmlData.Item(0).InnerText;
                string DecodingWorkHours = Encoding.Default.GetString(Convert.FromBase64String(workHoursBase64));
                XmlDocument workHoursXmlDoc = new XmlDocument();
                workHoursXmlDoc.LoadXml(DecodingWorkHours);

                string startTime = getXMLElementTextContentOrEmpty(workHoursXmlDoc.DocumentElement, "Start");
                string endTime = getXMLElementTextContentOrEmpty(workHoursXmlDoc.DocumentElement, "End");
                string bias = getXMLElementTextContentOrEmpty(workHoursXmlDoc.DocumentElement, "Bias");
                workDays = getXMLElementTextContentOrEmpty(workHoursXmlDoc.DocumentElement, "WorkDays");

                result = new WorkingHours
                {
                    StartDate = startTime == "" ? defaultStart : startTime,
                    EndDate = endTime ==  "" ? defaultEnd : endTime,
                    TimeZoneBias = bias == "" ? 0 : Convert.ToInt64(bias),
                    WorkDays = workDays == "" ? defaultWorkdays : parseWorkDays(workDays)
                };
            }
            
            //Console.WriteLine("-----------your startWork time:" + result.StartDate);
            //Console.WriteLine("-----------your endWork time:" + result.EndDate);
            //Console.WriteLine("-----------your timezone bias:" + result.TimeZoneBias);
            //Console.WriteLine("-----------your work days:" + workDays);
            
            return result;
        }

        public static Message[] replyToUserMailPropsParsed(string response)
        {
            Message[] result = { };
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(response);

            XmlNodeList messages = xmlDoc.GetElementsByTagName("t:Message");

            for (int i = 0; i < messages.Count; i++)
            {
                XmlElement conversationElement = (XmlElement)(getXMLExactlyOneElementOrThrow(messages.Item(i), "t:ConversationId"));
                string conversationId = conversationElement.GetAttribute("Id");

                string internetMsgId = getXMLExactlyOneElementOrThrow(messages.Item(i), "t:InternetMessageId").InnerText;

                string sentTimeTicks = getXMLExactlyOneElementOrThrow(messages.Item(i), "t:DateTimeSent").InnerText;
                DateTime itemSentTime = new DateTime();

                try
                {
                    itemSentTime = Convert.ToDateTime(sentTimeTicks.Substring(0, 10) + " " + sentTimeTicks.Substring(11, 8));
                }
                catch (Exception ex)
                {
                    //Console.WriteLine("Error: DateTimeSent is invalid when parsing response for GetItem...Skip this item...");
                    continue;
                }

                string inReplyTo = getXMLElementTextContentOrEmpty(messages.Item(i), "t:InReplyTo");

                XmlElement itemElement = (XmlElement)(getXMLExactlyOneElementOrThrow(messages.Item(i), "t:ItemId"));
                string itemId = itemElement.GetAttribute("Id");

                XmlNode mailboxNode = getXMLExactlyOneElementOrThrow(messages.Item(i), "t:Mailbox");
                string senderName = getXMLElementTextContentOrEmpty(mailboxNode, "t:Name");
                string senderAddress = getXMLElementTextContentOrEmpty(mailboxNode, "t:EmailAddress");
                if (senderName == "" || senderAddress == "")
                {
                    //Console.WriteLine("Error: senderName or senderAddress is empty when parsing response for GetItem...Skip this item:" + itemId);
                    continue;
                }
                string subject = getXMLElementTextContentOrEmpty(messages.Item(i), "t:Subject");

                string itemClass = getXMLElementTextContentOrEmpty(messages.Item(i), "t:ItemClass");

                //TODO: Meeting response ?????
                if(itemClass != "IPM.Note" && itemClass != "IPM.Schedule.Meeting.Request")
                {
                    continue;
                }
                Message sentMail = new Message
                {
                    InternetMsgId = internetMsgId,
                    ItemId = itemId,
                    ConversationId = conversationId,
                    SentDateTime = itemSentTime,
                    Subject = subject,
                    InReplyTo = inReplyTo,
                    Sender = new Person { Name = senderName, SmtpAddress = senderAddress, Title = "" }
                };

                result = result.Concat(new Message[] { sentMail }).ToArray();
            }

            return result;
        }

        public static int[] parseWorkDays(string workDays)
        {
            int[] result = { 0,0,0,0,0,0,0};
            foreach (string workDay in workDays.Split(' '))
            {
                result[workDayToNumber(workDay)] = 1;
            }
            return result;
        }

        private static int workDayToNumber(string workDay)
        {
            switch (workDay.ToLower())
            {
                case "sunday":
                    return 0;
                case "monday":
                    return 1;
                case "tuesday":
                    return 2;
                case "wednesday":
                    return 3;
                case "thursday":
                    return 4;
                case "friday":
                    return 5;
                case "saturday":
                    return 6;
                default:
                    return 0;
            }
        }
        public static string[] parseItems(XmlNode itemIdsXmlNode)
        {
            string[] result = { };
            XmlNodeList itemIdNodes = ((XmlElement)itemIdsXmlNode).GetElementsByTagName("ItemId");
            for (int i = 0; i < itemIdNodes.Count; i++)
            {
                XmlElement itemElement = (XmlElement)(itemIdNodes.Item(i));
                string itemId = itemElement.GetAttribute("Id");
                result = result.Concat(new string[] { itemId }).ToArray();
            }
            return result;
        }

        public static XmlNode getXMLExactlyOneElementOrThrow(XmlNode xmlElementRoot, string xmlTag)
        {
            XmlElement element = (XmlElement)(xmlElementRoot);
            XmlNode node = null;
            try
            {
                node = element.GetElementsByTagName(xmlTag).Item(0);
                
            }
            catch(Exception ex) {
               // Console.WriteLine("Can not find xmlTag: " + xmlTag + " in getXMLExactlyOneElementOrThrow");
            }
            return node;
        }

        public static string getXMLElementTextContentOrEmpty(XmlNode xmlElementRoot, string xmlTag)
        {
            try {
                XmlElement element = (XmlElement)(xmlElementRoot);
                XmlNodeList node = element.GetElementsByTagName(xmlTag);

                if (node.Count > 0)
                {
                    return node.Item(0).InnerText;
                }
            }catch (Exception ex)
            {
               // Console.WriteLine("Can not find xmlTag: " + xmlTag + " in getXMLElementTextContentOrEmpty");
            }

            return "";
        }
    }
}
