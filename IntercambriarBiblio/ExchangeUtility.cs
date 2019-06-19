using System;
using System.Net;
using Microsoft.Exchange.WebServices.Data;

namespace IntercambriarBiblio
{
    /// <summary>
    /// ExchangeUtility provides convenience methods for obtaining 
    /// EWS ExchangeService objects. This would only be useful if you needed 
    /// to use Exchange services not provided in the Calendar and Mail 
    /// namespaces. 
    /// <br/>
    /// For useful information on using Echange Web Services see:
    /// https://github.com/officedev/ews-managed-api
    /// and
    /// https://msdn.microsoft.com/en-us/library/office/dn567668(v=exchg.150).aspx
    /// </summary>
    public static class ExchangeUtility
    {
        static ExchangeUtility()
        {
            CertificateCallback.Initialize();
            ServicePointManager
                    .ServerCertificateValidationCallback +=
                (sender, cert, chain, sslPolicyErrors) => true;
        }

        // The following is a basic redirection validation callback method. It 
        // inspects the redirection URL and only allows the Service object to 
        // follow the redirection link if the URL is using HTTPS. 
        //
        // This redirection URL validation callback provides sufficient security
        // for development and testing of your application. However, it may not
        // provide sufficient security for your deployed application. You should
        // always make sure that the URL validation callback method that you use
        // meets the security requirements of your organization.
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            var result = false;

            var redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }

            return result;
        }

        /// <summary>
        /// Simple connect to service method.
        /// </summary>
        /// <param name="userData"></param>
        /// <returns></returns>
        public static ExchangeService ConnectToService(IExchangeUserData userData)
        {
            return ConnectToService(userData, null);
        }

        public static ExchangeService ConnectToService(IExchangeUserData userData, ITraceListener listener)
        {
            var service = new ExchangeService(userData.Version);

            if (listener != null)
            {
                service.TraceListener = listener;
                service.TraceFlags = TraceFlags.All;
                service.TraceEnabled = true;
            }

            service.Credentials = new NetworkCredential(userData.EmailAddress, userData.Password);

            if (userData.AutodiscoverUrl == null)
            {
                //Console.Write($"Using Autodiscover to find EWS URL for {userData.EmailAddress}. Please wait... ");

                service.AutodiscoverUrl(userData.EmailAddress, RedirectionUrlValidationCallback);
                userData.AutodiscoverUrl = service.Url;

                //Console.WriteLine("Complete");
            }
            else
            {
                service.Url = userData.AutodiscoverUrl;
            }

            return service;
        }

        public static ExchangeService ConnectToServiceWithImpersonation(
            IExchangeUserData userData,
            string impersonatedUserSmtpAddress)
        {
            return ConnectToServiceWithImpersonation(userData, impersonatedUserSmtpAddress, null);
        }

        public static ExchangeService ConnectToServiceWithImpersonation(
            IExchangeUserData userData,
            string impersonatedUserSmtpAddress,
            ITraceListener listener)
        {
            var service = new ExchangeService(userData.Version);

            if (listener != null)
            {
                service.TraceListener = listener;
                service.TraceFlags = TraceFlags.All;
                service.TraceEnabled = true;
            }

            service.Credentials = new NetworkCredential(userData.EmailAddress, userData.Password);

            var impersonatedUserId =
                new ImpersonatedUserId(ConnectingIdType.SmtpAddress, impersonatedUserSmtpAddress);

            service.ImpersonatedUserId = impersonatedUserId;

            if (userData.AutodiscoverUrl == null)
            {
                service.AutodiscoverUrl(userData.EmailAddress, RedirectionUrlValidationCallback);
                userData.AutodiscoverUrl = service.Url;
            }
            else
            {
                service.Url = userData.AutodiscoverUrl;
            }

            return service;
        }
    }
}

