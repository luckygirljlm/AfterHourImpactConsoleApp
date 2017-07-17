using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AfterHourConsoleApplication
{
    public class GetUserConfigurationRequest : EWSRequestBase
    {
        public override string[] additionalFieldURIProperties()
        {
            string[] properties = {};
            return properties;
        }
        public override string getSOAPRequest()
        {
            return @"<m:GetUserConfiguration>
                <m:UserConfigurationName Name=""WorkHours"">
                    <t:DistinguishedFolderId Id=""calendar""/>
                </m:UserConfigurationName>
                <m:UserConfigurationProperties>XmlData</m:UserConfigurationProperties>
            </m:GetUserConfiguration>";
        }
    }
}
