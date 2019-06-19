using System;
using System.Security;
using Microsoft.Exchange.WebServices.Data;

namespace IntercambriarBiblio
{
    public class ExchangeUser : ExchangePrinciple, IExchangeUserData
    {

        public ExchangeUser(string emailAddress, SecureString password) : base(emailAddress)
        {
            Password = password;
            Version = ExchangeVersion.Exchange2010_SP2;
        }

        public ExchangeUser(string emailAddress, string password) : base(emailAddress)
        {
            SetPassword(password);
            Version = ExchangeVersion.Exchange2010_SP2;
        }

        private void SetPassword(string pw)
        {
            var sString = new SecureString();
            // Use the AppendChar method to add each char value to the secure string.
            foreach (char ch in pw)
                sString.AppendChar(ch);
            Password = sString;            
        }

        public ExchangeVersion Version { get; set; }
        public string EmailAddress => Id;
        public SecureString Password { get; set; }
        public Uri AutodiscoverUrl { get; set; }
    }
}
