using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AfterHourConsoleApplication
{
    public abstract class EWSRequestBase
    {
        public abstract string[] additionalFieldURIProperties();
        public abstract string getSOAPRequest();
        public string getSoapEnvelope() {
            string soapRequest =
            @"<?xml version=""1.0"" encoding=""utf-8""?>
             <soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
                           xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
                           xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
                           xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types""
                            xmlns:m=""http://schemas.microsoft.com/exchange/services/2006/messages"">
                <soap:Header>
                    <t:RequestServerVersion Version= ""Exchange2013"" />
                </soap:Header>
                <soap:Body>" +
                this.getSOAPRequest() +
                @"</soap:Body>
            </soap:Envelope>";

            return soapRequest;
        }

        protected string buildFieldURIPropertiesElements()
        {
            string result = "";
            foreach (string propertyURI in this.additionalFieldURIProperties())
            {
                string prop = @"<t:FieldURI FieldURI = """ + propertyURI + @""" />";
                if (result != "")
                {
                    result += " ";
                }
                result += prop;
            }
            return result;
        }
    }    
}
