using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AfterHourConsoleApplication
{
    public class FindUserSentItemsRequest : EWSRequestBase
    {
        private DateTime startDateOfQueryRange = Utils.generatePriorDayBreak(7);
        private DateTime endDateOfQueryRange = Utils.generateCurrentDayBreak();

        public FindUserSentItemsRequest(DateTime start)
        {
            this.startDateOfQueryRange = start;
        }

        public FindUserSentItemsRequest() { }

        public override string[] additionalFieldURIProperties()
        {
            string[] properties = {
                     "message:InternetMessageId",
                     "item:DateTimeSent",
                     "item:Subject",
                     "item:ConversationId",
                     "item:ItemId",
                     "item:InReplyTo",
                     "message:Sender"};
            return properties;
        }

        public override string getSOAPRequest()
        {
            return @"<m:FindItem Traversal=""Shallow"">
                    <m:ItemShape>
                        <t:BaseShape>IdOnly</t:BaseShape>
                        <t:AdditionalProperties>" +
                            this.buildFieldURIPropertiesElements() +
                        @"</t:AdditionalProperties>
                    </m:ItemShape>" +
                    this.buildRestriction() +
                    @"<m:ParentFolderIds>
                        <t:DistinguishedFolderId Id=""sentitems"" />
                    </m:ParentFolderIds>
                </m:FindItem>";
        }

        string buildRestriction()
        {
            string restriction =
                @"<m:Restriction>
                        <t:And>
                            <t:IsGreaterThan>
                                <t:FieldURI FieldURI=""item:DateTimeSent""/>
                                <t:FieldURIOrConstant>
                                    <t:Constant Value=""" + Utils.toISOString(startDateOfQueryRange) + @""" />" +
                                    @"</t:FieldURIOrConstant>
                            </t:IsGreaterThan>
                            <t:IsLessThan>
                                <t:FieldURI FieldURI=""item:DateTimeSent""/>
                                <t:FieldURIOrConstant>
                                    <t:Constant Value=""" + Utils.toISOString(endDateOfQueryRange) + @""" />" +
                                    @"</t:FieldURIOrConstant>
                            </t:IsLessThan>
                            <t:Or>
                                <t:IsEqualTo>
                                    <t:FieldURI FieldURI=""item:ItemClass""></t:FieldURI>
                                    <t:FieldURIOrConstant>
                                        <t:Constant Value=""IPM.Note""></t:Constant>
                                    </t:FieldURIOrConstant>
                                </t:IsEqualTo>
                                <t:IsEqualTo>
                                    <t:FieldURI FieldURI=""item:ItemClass""></t:FieldURI>
                                    <t:FieldURIOrConstant>
                                        <t:Constant Value=""IPM.Schedule.Meeting.Request""></t:Constant>
                                    </t:FieldURIOrConstant>
                                </t:IsEqualTo>
                            </t:Or>
                        </t:And>
                    </m:Restriction>";

            return restriction;
        }
    }
}
