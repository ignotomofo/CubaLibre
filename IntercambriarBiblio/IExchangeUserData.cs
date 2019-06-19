using System;
using System.Security;
using Microsoft.Exchange.WebServices.Data;

namespace IntercambriarBiblio
{
    /// <summary>
    /// Data elements required for an EWS user. 
    /// </summary>
    public interface IExchangeUserData
    {
        ExchangeVersion Version { get; }
        string EmailAddress { get; }
        SecureString Password { get; }
        Uri AutodiscoverUrl { get; set; }
    }
}