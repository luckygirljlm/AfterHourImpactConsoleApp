using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AfterHourConsoleApplication
{
    public class GetItemRequest : EWSRequestBase
    {
        private string[] itemIds = { };

        public GetItemRequest(string[] itemIds)
        {
            this.itemIds = itemIds;
        }
        public override string[] additionalFieldURIProperties()
        {
            string[] properties = {
                     "item:ItemId",
                     "item:ItemClass",
                     "item:DateTimeSent",
                     "message:InternetMessageId",
                     "item:ConversationId",
                     "item:Subject",
                     "item:InReplyTo",
                     "message:Sender" };
            return properties;
        }

        public override string getSOAPRequest()
        {
            return @"<m:GetItem>
                <m:ItemShape>
                    <t:BodyType>Text</t:BodyType>
                    <t:BaseShape>IdOnly</t:BaseShape>
                    <t:AdditionalProperties>" +
                        buildFieldURIPropertiesElements() +
                    @"</t:AdditionalProperties>
                </m:ItemShape>
                <m:ItemIds>" +
                    buildItemIdElements() +
                @"</m:ItemIds>
            </m:GetItem>";
        }

        public string buildItemIdElements()
        {
            string result = "";
            foreach (string itemId in itemIds)
            {
                string itemTag = @"<t:ItemId Id= """ + itemId + @""" />";
                if (result != "")
                {
                    result += " ";
                }
                result += itemTag;
            }
            return result;
        }
    }
}
