using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Exchange;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Runtime;
using System.Xml;
using System.IO;
using AfterHourConsoleApplication.SDK;

namespace AfterHourConsoleApplication
{
    static class AfterHourImpactComputor
    {
        public static void getAfterHourImpact()
        {
            Message[] sentMails = getEmailsSentAfterHour();
            SDK.Conversation[] joinedConvs = getConversationsUserJoined();
            SDK.Conversation[] afterHourJoinedConvs = getValidConversations(sentMails, joinedConvs);

            //logConversationTopic(afterHourJoinedConvs);

            Message[] detailedConvMails = getDetailedConvs(afterHourJoinedConvs);
            
            //TODO: call weve api to get all recipients' working hour

            Dictionary<string, List<ReplyMessage>> statistics = getStatistics(sentMails, detailedConvMails);
            generateStatisticsReport(statistics);
        }

        static Message[] getEmailsSentAfterHour()
        {
            //Console.WriteLine("///////////////getting your sentEmails in After hour...");

            string sentMailsResponse = ExecuteRequest(new FindUserSentItemsRequest(Utils.startRange), "UserSentItems", true);
            Message[] sentMails = ParseUtils.sentMailPropsParsed(sentMailsResponse);
            Message[] sentMailsInAfterHour = filterAfterHour(sentMails);

            Console.WriteLine("you sent " + sentMailsInAfterHour.Length + " emails in your after hour " + Utils.lastServeralDaysText);

            Utils.addToProcessReport("you sent " + sentMailsInAfterHour.Length.ToString() + " emails in your after hour " + Utils.lastServeralDaysText + "\r\n\r\n");

            //logSentMails(sentMailsInAfterHour);
            //Console.WriteLine("///////////////End...\n\n");
            return sentMailsInAfterHour;
        }

        static SDK.Conversation[] getConversationsUserJoined()
        {
            //Console.WriteLine("///////////////getting conversations you joined...");
            string joinedConvResponse = ExecuteRequest(new FindUserJoinedConversationsRequest(Utils.startRange), "UserJoinedConversations", true);
            SDK.Conversation[] joinedConversations = ParseUtils.joinedConversationPropsParsed(joinedConvResponse);
            //Console.WriteLine("you joined " + joinedConversations.Length + " conversations " + Utils.lastServeralDaysText);
            //Console.WriteLine("///////////////End...\n\n");
            return joinedConversations;
        }

        static List<ReplyMessage> convertToReplyMessage(List<Message> mails)
        {
            List<ReplyMessage> result = new List<ReplyMessage>();
            foreach (Message mail in mails)
            {
                Message inReplyToMail = findInReplyToMailByInternetMsgId(mails.ToArray(), mail.InReplyTo);

                result.Add(new ReplyMessage
                {
                    InternetMsgId = mail.InternetMsgId,
                    ItemId = mail.ItemId,
                    ConversationId = mail.ConversationId,
                    SentDateTime = mail.SentDateTime,
                    Subject = mail.Subject,
                    InReplyTo = mail.InReplyTo,
                    Sender = mail.Sender,
                    InReplyToItemId = inReplyToMail != null ? inReplyToMail.ItemId : "",
                    InReplyToSentDateTime = inReplyToMail != null ? inReplyToMail.SentDateTime : DateTime.Now.AddDays(-100)
                });
            }

            return result;
        }

        static Message[] filterAfterHour(Message[] mails)
        {
            Message[] mailsInAfterHour = { };

            foreach (Message email in mails)
            {
                if (Utils.isAfterHour(email.SentDateTime, Utils.userWorkingHourConfig))
                {
                    mailsInAfterHour = mailsInAfterHour.Concat(new Message[] { email }).ToArray();
                }
            }
            return mailsInAfterHour;
        }

        public static string ExecuteRequest(EWSRequestBase ewsRequest, string fileName, bool shouldLogTime)
        {
            HttpWebRequest request = CreateWebRequest();

            XmlDocument soapEnvelopeXml = new XmlDocument();
            soapEnvelopeXml.LoadXml(ewsRequest.getSoapEnvelope());

            //writeToFile(ewsRequest.getSoapEnvelope(), fileName + "_request.xml");

            DateTime start = DateTime.Now;

            using (Stream stream = request.GetRequestStream())
            {
                soapEnvelopeXml.Save(stream);
            }

            using (WebResponse response = request.GetResponse())
            {
                using (StreamReader rd = new StreamReader(response.GetResponseStream()))
                {
                    string soapResult = rd.ReadToEnd();
                    //writeToFile(soapResult, fileName + "_result.xml");
                    TimeSpan span = DateTime.Now - start;

                    if (shouldLogTime)
                        Console.WriteLine("\n========================execute " + fileName + " Request take " + (span.Seconds * 1000 + span.Milliseconds) + "ms\n");

                    Utils.addToProcessReport("\r\n===============================execute " + fileName + " Request take " + (span.Seconds * 1000 + span.Milliseconds).ToString() + "ms\r\n\r\n");

                    return soapResult;
                }
            }

        }

        public static void writeToFile(string content, string fileName)
        {
            string filePath = Utils.getFilePath();
            if (!Directory.Exists(filePath))
            {
                DirectoryInfo di = Directory.CreateDirectory(filePath);
            }

            FileStream fileStream = new FileStream(filePath + fileName, FileMode.Create, FileAccess.Write);
            StreamWriter sWriter = new StreamWriter(fileStream);
            sWriter.Write(content);
            sWriter.Close();
        }

        public static HttpWebRequest CreateWebRequest()
        {
            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(@"https://outlook.office365.com/ews/exchange.asmx");
            webRequest.Headers.Add(@"SOAP:Action");
            webRequest.Headers.Add("Authorization", "Bearer " + Utils.token);
            webRequest.ContentType = "text/xml;charset=\"utf-8\"";
            webRequest.Accept = "text/xml";
            webRequest.Method = "POST";
            return webRequest;
        }

        static SDK.Conversation[] getValidConversations(Message[] sentMails, SDK.Conversation[] joinedConvs)
        {
            if (Utils.shouldStartConv)
            {
                Message[] startMails = getMailsUserStarted(sentMails);
                return getAfterHourStartAndRepliedConversations(startMails, joinedConvs);  // after hour start and has replies
            }
            else {
                return getJoinedAfterHourConversations(sentMails, joinedConvs);  // after hour joined and has replies
            }
        }

        static Message[] getDetailedConvs(SDK.Conversation[] afterHourJoinedConvs)
        {
            string[] itemIdsToQuery = getItemIdsToQuery(afterHourJoinedConvs);

            if (itemIdsToQuery.Length == 0)
            {
                return new Message[] { };
            }
            
            //call GetItems
            string replyMailsResponse = ExecuteRequest(new GetItemRequest(itemIdsToQuery), "getDetailedConvs", true);
            Message[] detailedMailsInConv = ParseUtils.replyToUserMailPropsParsed(replyMailsResponse);  // filter auto reply

            return detailedMailsInConv;
        }

        //return recipients -> replyMessages
        static Dictionary<string, List<ReplyMessage>> getStatistics(Message[] sentMails, Message[] detailedMailsInConv)
        {
            //group by convid and order by sentTime
            Dictionary<string, List<Message>> detailedConvDictionary = getMessageGroupByConversation(detailedMailsInConv);  // after hour joined conversation
            Dictionary<string, List<Message>> sentMailsDictionary = getMessageGroupByConversation(sentMails);  // after hour sent mails

            Dictionary<string, List<ReplyMessage>> convertedDetailedConvDictionary = new Dictionary<string, List<ReplyMessage>>();
            Dictionary<string, List<ReplyMessage>> result = new Dictionary<string, List<ReplyMessage>>();

            foreach (var convId in detailedConvDictionary.Keys)
            {
                if (sentMailsDictionary.ContainsKey(convId)) {
                    List<ReplyMessage> replyValue = convertToReplyMessage(detailedConvDictionary[convId]);
                    convertedDetailedConvDictionary.Add(convId, replyValue);

                    Dictionary<string, List<ReplyMessage>> currentImpact = new Dictionary<string, List<ReplyMessage>>();
                    if (Utils.shouldDirectReply)
                    {
                        currentImpact = getImpactToDirectReplies(replyValue, sentMailsDictionary[convId]);
                    }
                    else
                    {
                        currentImpact = getImpactToAllReplies(replyValue, sentMailsDictionary[convId]);
                    }
                    mergeDictionary(result, currentImpact);
                }
            }
            return result;
        }

        static void mergeDictionary(Dictionary<string, List<ReplyMessage>> target, Dictionary<string, List<ReplyMessage>> source)
        {
            foreach (string key in source.Keys)
            {
                if (target.ContainsKey(key))
                {
                    target[key] = target[key].Concat(source[key]).ToList();
                }
                else
                {
                    target.Add(key, source[key]);
                }
            }
        }

        static Dictionary<string, List<ReplyMessage>> getImpactToDirectReplies(List<ReplyMessage> convMessages, List<Message> afterHourSentMessages)
        {
            Dictionary<string, List<ReplyMessage>> impact = new Dictionary<string, List<ReplyMessage>>();
            List<Message> sentMessagesToComputeImpact = getAfterHourMessagesToComputeImpact(convMessages, afterHourSentMessages);
            foreach (ReplyMessage mail in convMessages)
            {
                string impactedUserName = mail.Sender.Name;
                // filter user itself
                if (mail.InReplyToItemId != "" && impactedUserName != Utils.senderName &&
                    existItem(sentMessagesToComputeImpact.ToArray(), mail.InReplyToItemId))
                {
                    SDK.WorkingHours wk = getWorkingHours(impactedUserName);
                    if (Utils.isAfterHourReceivedAndReply(mail.SentDateTime, mail.InReplyToSentDateTime, wk))
                    {
                        if (!impact.ContainsKey(impactedUserName))
                        {
                            impact.Add(impactedUserName, new List<ReplyMessage>());
                        }
                        impact[impactedUserName].Add(mail);
                    }
                }
            }
            return impact;
        }

        static ReplyMessage getLastAfterHourMailFromOthers(List<ReplyMessage> convMessages, Message firstAfterHourMailFromUser)
        {
            ReplyMessage result = null;
            for (int i = 0; i < convMessages.Count; i++)
            {
                string senderName = convMessages[i].Sender.Name;
                if (convMessages[i].ItemId == firstAfterHourMailFromUser.ItemId)
                    break;
                if (Utils.isAfterHour(convMessages[i].SentDateTime, getWorkingHours(senderName)))
                    result = convMessages[i];
            }
            return result;
        }

        static List<Message> getAfterHourMessagesToComputeImpact(List<ReplyMessage> convMessages, List<Message> afterHourMessages)
        {
            ReplyMessage lastAfterHourMailFromOthers = getLastAfterHourMailFromOthers(convMessages, afterHourMessages[0]);
            if (lastAfterHourMailFromOthers == null)
                return afterHourMessages;
            List<Message> result = new List<Message>(afterHourMessages);

            for (int i = 0; i < afterHourMessages.Count; i++)
            {
                ReplyMessage mail = getReplyMessageFromConvByItemId(convMessages, afterHourMessages[i].ItemId);
                if (mail != null)
                {
                    if (mail.InReplyToSentDateTime < Utils.getLastEndWorkHour(mail.SentDateTime, Utils.userWorkingHourConfig)
                        || lastAfterHourMailFromOthers.SentDateTime.AddDays(1) < mail.SentDateTime)
                        break;
                }
                
                result.RemoveAt(0);
                
            }
            return result;
        }

        static ReplyMessage getReplyMessageFromConvByItemId(List<ReplyMessage> convMessages, string itemId)
        {
            foreach (ReplyMessage mail in convMessages)
            {
                if (mail.ItemId == itemId)
                    return mail;
            }
            return null;
        }

        static List<List<ReplyMessage>> splitConvMessagesBySentAfterHour(List<ReplyMessage> convMessages, List<Message> afterHourMessages)
        {
            List<List<ReplyMessage>> result = new List<List<ReplyMessage>>();
            int afterHourSentIndex = 0;
            bool stopImpact = false;
            List<ReplyMessage> currentList = new List<ReplyMessage>();
            for (var i = 0; i < convMessages.Count; i++)
            {
                ReplyMessage mail = convMessages[i];
                if (afterHourSentIndex < afterHourMessages.Count && mail.ItemId == afterHourMessages[afterHourSentIndex].ItemId) // user sent in after hour
                {
                    if (currentList.Count > 0)
                    {
                        result.Add(currentList);
                        currentList = new List<ReplyMessage>();
                    }

                    currentList.Add(mail);
                    afterHourSentIndex++;
                    stopImpact = false;
                }
                else if (mail.Sender.Name == Utils.senderName && afterHourSentIndex > 0)  // you sent emails between your 2 after hr email, reduce the former impact
                {
                    stopImpact = true;
                }
                else if (!stopImpact &&
                    mail.Sender.Name != Utils.senderName && 
                    Utils.isAfterHour(mail.SentDateTime, getWorkingHours(mail.Sender.Name)))  // others sent in their after hour
                {
                    if (afterHourSentIndex == 0) // first batch
                    {
                        currentList.Add(mail);
                        continue;
                    }
                    else if (Utils.isAfterHourReceivedAndReply(mail.SentDateTime, mail.InReplyToSentDateTime, getWorkingHours(mail.Sender.Name)))  //others may impacted by user
                    {
                        currentList.Add(mail);
                    }
                }
            }
            if(currentList.Count > 0)
            {
                result.Add(currentList);
            }
            return result;
        }
        static Dictionary<string, List<ReplyMessage>> getImpactToAllReplies(List<ReplyMessage> convMessages, List<Message> afterHourMessages)
        {
            Dictionary<string, List<ReplyMessage>> impact = new Dictionary<string, List<ReplyMessage>>();
            List<List<ReplyMessage>> impactMailsBatch = splitConvMessagesBySentAfterHour(convMessages, afterHourMessages);

            bool isUserFirstAfterHour = impactMailsBatch[0][0].ItemId == afterHourMessages[0].ItemId;
            int impactStartIndex = isUserFirstAfterHour ? 0 : -1;

            for(var i = 0 ; i < impactMailsBatch.Count; i++)
            {
                Dictionary<string, List<ReplyMessage>> currentImpact = new Dictionary<string, List<ReplyMessage>>();
                if(impactStartIndex >= 0 && i >= impactStartIndex)
                {
                    currentImpact = calculateImpact(impactMailsBatch[i]);  
                }
                else if(impactStartIndex == -1 && i > 0)// user is not the first one
                {
                    ReplyMessage lastOtherWorkingAfterHour = impactMailsBatch[0][impactMailsBatch[0].Count - 1];
                    // there is workhour block between receivedTime and sentTime or 24hrs passed since last sent after hour by others
                    if (impactMailsBatch[i][0].InReplyToSentDateTime < Utils.getLastEndWorkHour(impactMailsBatch[i][0].SentDateTime, Utils.userWorkingHourConfig)
                        || lastOtherWorkingAfterHour.SentDateTime.AddDays(1) < impactMailsBatch[i][0].SentDateTime)
                    {
                        impactStartIndex = i;
                        currentImpact = calculateImpact(impactMailsBatch[i]);
                    }
                }

                mergeDictionary(impact, currentImpact);
            }
            return impact;
        }

        static DateTime getImpactDateBoundary(DateTime date)  // ensure date is in after hour
        {
            DateTime nextStartWorkDate = Utils.getNextStartWorkHour(date, Utils.userWorkingHourConfig);
            DateTime nextDay = date.AddDays(1);
            return nextStartWorkDate > nextDay ? nextStartWorkDate : nextDay;
        }

        static Dictionary<string, List<ReplyMessage>> calculateImpact(List<ReplyMessage> messagesReceivedAndRepliedAfterHour)
        {
            Dictionary<string, List<ReplyMessage>> result = new Dictionary<string, List<ReplyMessage>>();
            DateTime dateTimeBoundary = getImpactDateBoundary(messagesReceivedAndRepliedAfterHour[0].SentDateTime);

            foreach (ReplyMessage mail in messagesReceivedAndRepliedAfterHour)
            {
                string name = mail.Sender.Name;
                if (name != Utils.senderName && mail.SentDateTime < dateTimeBoundary)
                {
                    if (!result.ContainsKey(name))
                    {
                        result.Add(name, new List<ReplyMessage>());
                    }
                    result[name].Add(mail);
                }
            }
            return result;
        }

        static SDK.WorkingHours getWorkingHours(string name)
        {
            name = name.ToLower().Replace(" ", "").Trim();
            if (Utils.receipientsConfig.ContainsKey(name))
            {
                return Utils.receipientsConfig[name];
            }
            else
            {
                return Utils.defaultWorkingHours;
            }
        }
        static Dictionary<string, List<Message>> getMessageGroupByConversation(Message[] mails)
        {
            Dictionary<string, List<Message>> result = new Dictionary<string, List<Message>>();
            foreach (var mail in mails)
            {
                if (result.ContainsKey(mail.ConversationId))
                {
                    result[mail.ConversationId].Add(mail);
                }
                else
                {
                    result.Add(mail.ConversationId, new List<Message> { mail });
                }
            }

            Dictionary<string, List<Message>> orderedResult = new Dictionary<string, List<Message>>();
            //order by sentTime asc
            foreach (var key in result.Keys)
            {
                List<Message> value = result[key];
                orderedResult.Add(key, value.OrderBy(x => x.SentDateTime).ToList());
            }

            return orderedResult;
        }

        static string[] getItemIdsToQuery(SDK.Conversation[] afterHourJoinedConvs)
        {
            string[] itemIdsToQuery = new string[] { };

            foreach (var conv in afterHourJoinedConvs)
            {
                itemIdsToQuery = itemIdsToQuery.Concat(conv.ItemsAll).ToArray();
            }

            return itemIdsToQuery;
        }

        static void logConversationTopic(SDK.Conversation[] convs)
        {
            if (convs.Length > 0)
            {
                Console.WriteLine("ConversationTopics:");
                Utils.addToProcessReport("ConversationTopics:\r\n");

                foreach (SDK.Conversation conv in convs)
                {
                    Console.WriteLine(conv.ConversationTopic);
                    Utils.addToProcessReport(conv.ConversationTopic + "\r\n");
                }
            }
        }

        static void logSentMails(Message[] messages)
        {
            if (messages.Length > 0)
            {
                Console.WriteLine("Mails Info:");
                foreach (Message mail in messages)
                {
                    //string text = "Subject:" + mail.Subject + "  Sent on:" + mail.SentDateTime + " In UTC";
                    string text = "Subject:" + mail.Subject + "  Sent on:" + Utils.toLocaleTime(mail.SentDateTime, Utils.userWorkingHourConfig);
                    Console.WriteLine(text);
                    Utils.addToProcessReport(text + "\r\n");
                }
                Utils.addToProcessReport("\r\n\r\n\r\n");
            }
        }

        static void generateStatisticsReport(Dictionary<string, List<ReplyMessage>> report)
        {
            //Console.
            string wkConfig = "\r\n\r\nYour Working Hours Config:\r\n" + UserConfig.printWorkingHour(Utils.userWorkingHourConfig);

            Console.WriteLine("\n\n////////////AfterHourImpact Report");
            Console.WriteLine("ImpactedReceipient\tEmailCount");

            int receipientsCount = 0;
            string reportString = wkConfig;
            foreach (KeyValuePair<string, List<ReplyMessage>> kvp in report)
            {
                receipientsCount++;
                Console.WriteLine(kvp.Key + "\t\t" + kvp.Value.Count);

                reportString = reportString + "\r\n==========From " + kvp.Key + ":\r\n";
                int index = 1;
                foreach (ReplyMessage mail in kvp.Value)
                {
                    reportString = reportString + index.ToString() + "\r\n";
                    string itemInfo = "";
                    itemInfo = itemInfo + "Subject:" + mail.Subject + "\r\n";
                    itemInfo = itemInfo + "itemId:" + mail.ItemId + "\r\n";
                    itemInfo = itemInfo + "sentTime:" + mail.SentDateTime + "\r\n";
                    itemInfo = itemInfo + "inReplyToMailItemId:" + mail.InReplyToItemId + "\r\n";
                    itemInfo = itemInfo + "inReplyToMailSentTime:" + mail.InReplyToSentDateTime + "\r\n";
                    reportString = reportString + itemInfo;
                    index++;
                }
            }

            string conclusion = "\r\n\r\nYou have an after hour impact on " + receipientsCount.ToString() + " persons in total.";
            reportString = reportString + conclusion;
            string fileName = Utils.resultFileName + "_Start_" + Utils.shouldStartConv + "_DirectReply_" + Utils.shouldDirectReply + ".txt";
            writeToFile(reportString, fileName);

            Console.WriteLine(reportString);
            Console.WriteLine(string.Format("\n\nReports are saved to \"{0}\"...", Utils.filePath));
        }

        static Message[] getMailsUserStarted(Message[] sentMails)
        {
            Message[] result = { };
            foreach (Message email in sentMails)
            {
                if (email.InReplyTo == "")
                {
                    result = result.Concat(new Message[] { email }).ToArray();
                }
            }

            return result;
        }

        static SDK.Conversation[] getJoinedAfterHourConversations(Message[] sentMails, SDK.Conversation[] joinedConvs)
        {
            SDK.Conversation[] result = { };
            foreach (SDK.Conversation conv in joinedConvs)
            {
                if (!hasReply(conv))
                    continue;

                List<string> sentItemIds = new List<string>();
                foreach (var message in sentMails) {
                    sentItemIds.Add(message.ItemId);
                }

                List<string> convSentItemIds = new List<string>();
                convSentItemIds = new List<string>(conv.ItemsSent);
                
                List<string> interSect = convSentItemIds.Intersect(sentItemIds).ToList();
                if (interSect.Count > 0) {
                    result = result.Concat(new SDK.Conversation[] { conv }).ToArray();
                }
            }

            return result;
        }

        static bool hasReply(SDK.Conversation conv)
        {
            bool hasReplies = conv.ItemsAll.Length > conv.ItemsSent.Length;
            if (!hasReplies)
                return false;
            if (conv.ItemsSent.Length == 1 && conv.ItemsAll[0] == conv.ItemsSent[0])  // simply filter one sent but no reply
                return false;
            return true;
        }

        static SDK.Conversation[] getAfterHourStartAndRepliedConversations(Message[] startMails, SDK.Conversation[] joinedConvs)
        {
            SDK.Conversation[] result = { };
            foreach (SDK.Conversation conv in joinedConvs)
            {
                string startItemId = conv.ItemsAll[conv.ItemsAll.Length - 1];
                if (!hasReply(conv))
                    continue;

                bool startInAfterHour = false;
                foreach (Message email in startMails)
                {
                    if (email.ConversationId == conv.ConversationId && email.ItemId == startItemId)
                    {
                        startInAfterHour = true;
                        break;
                    }
                }
                if (startInAfterHour)
                {
                    result = result.Concat(new SDK.Conversation[] { conv }).ToArray();
                }
            }

            return result;
        }

        static Message[] getSentMailsToComputeImpact(Message[] sentMails, SDK.Conversation[] startAndRepliedConvs)
        {
            Message[] result = { };
            foreach (Message email in sentMails)
            {
                bool hasImpact = false;
                foreach (SDK.Conversation conv in startAndRepliedConvs)
                {
                    if (conv.ConversationId == email.ConversationId)
                    {
                        hasImpact = true;
                        break;
                    }
                }
                if (hasImpact)
                {
                    result = result.Concat(new Message[] { email }).ToArray();
                }
            }

            return result;
        }

        static bool existItem(Message[] mails, string itemId)
        {
            foreach (Message mail in mails)
            {
                if (mail.ItemId == itemId)
                {
                    return true;
                }
            }
            return false;
        }

        static Message findInReplyToMailByInternetMsgId(Message[] mails, string internetMsgId)
        {
            if (internetMsgId == "")
                return null;
            foreach (Message mail in mails)
            {
                if (mail.InternetMsgId == internetMsgId)
                {
                    return mail;
                }
            }
            return null;
        }
    }
}
