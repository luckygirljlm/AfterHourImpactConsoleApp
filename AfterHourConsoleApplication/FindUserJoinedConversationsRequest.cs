using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AfterHourConsoleApplication
{
    public class FindUserJoinedConversationsRequest : EWSRequestBase
    {
        private DateTime startDateOfQueryRange = Utils.generatePriorDayBreak(7);
        private DateTime endDateOfQueryRange = Utils.generatePriorDayBreak(-1);

        public FindUserJoinedConversationsRequest(DateTime start)
        {
            this.startDateOfQueryRange = start;
        }

        public FindUserJoinedConversationsRequest() { }
        public override string[] additionalFieldURIProperties()
        {
            string[] properties = {
                "conversation:ConversationId",
                "conversation:ConversationTopic",
                "conversation:ItemIds",
                "conversation:GlobalItemIds" };
            return properties;
        }

        public override string getSOAPRequest()
        {
            return @"<m:FindConversation Traversal=""Deep"" ViewFilter=""All"">
                <m:QueryString>" + this.buildQueryString() + @"</m:QueryString>          
                <m:ParentFolderId>
                    <t:DistinguishedFolderId Id=""sentitems""/>
                </m:ParentFolderId>
                <m:ConversationShape>
                    <t:BaseShape>IdOnly</t:BaseShape>
                    <t:AdditionalProperties>" +
                        this.buildFieldURIPropertiesElements() +
                    @"</t:AdditionalProperties>
                </m:ConversationShape>
            </m:FindConversation>";
        }

        string buildQueryString() {
            string query=@"Sent:" + Utils.getDateInSlashFormat(this.startDateOfQueryRange) + ".." + Utils.getDateInSlashFormat(this.endDateOfQueryRange);
            return query;
        }
    }
}
