using System;
using System.IO;
using Microsoft.Exchange.WebServices.Data;

namespace IntercambriarBiblio.Mail
{
    public class MailService
    {
        private readonly IExchangeUserData _user;
        private ExchangeService _service;
        private ExchangeService Service => _service ?? (_service = GetExchangeService());

        public MailService(string email, string password)
        {
            _user = new ExchangeUser(email, password);
        }

        public MailService(string email, string password, Uri service)
        {
            _user = new ExchangeUser(email, password) {AutodiscoverUrl = service};
        }

        public MailService(IExchangeUserData userData)
        {
            _user = userData;
        }

        private ExchangeService GetExchangeService()
        {
            return ExchangeUtility.ConnectToService(_user);
        }

        /// <summary>
        /// Attempt to send an email through exchange. The user logged
        /// into the service is the sender. 
        /// </summary>
        /// <param name="mail">The email message to send.</param>
        public void SendMail(Email mail)
        {
            var email = new EmailMessage(Service);

            foreach (var r in mail.ToRecipients)
            {
                email.ToRecipients.Add(r.Id);
            }
            foreach (var r in mail.CcRecipients)
            {
                email.CcRecipients.Add(r.Id);
            }
            foreach (var r in mail.BccRecipients)
            {
                email.BccRecipients.Add(r.Id);
            }            
            email.Subject = mail.Subject;
            var bt = (mail.BodyType == EmailBodyType.Text) ? BodyType.Text : BodyType.HTML;
            email.Body = new MessageBody(bt, mail.Body);                        
            email.Send();
        }

        public void ExportMail(string exportFolder, WellKnownFolderName wkfn)
        {
            DirectoryInfo di;
            if (!Directory.Exists(exportFolder))
            {
                di = Directory.CreateDirectory(exportFolder);
            }
            else
            {
                di = new DirectoryInfo(exportFolder);
            }

            ExportMimeEmail(Service, wkfn, exportFolder);
        }

        private static void ExportMimeEmail(ExchangeService service, WellKnownFolderName folderName, string path)
        {
            const int offset = 0;
            const int pageSize = 100;
            var more = true;
            Console.WriteLine($"URL: {service.Url}");
            var inbox = Folder.Bind(service, folderName);
            var view = new ItemView(pageSize, offset, OffsetBasePoint.Beginning)
            {
                PropertySet = new PropertySet(ItemSchema.Id, ItemSchema.DateTimeCreated, ItemSchema.DateTimeReceived)
            };

            while (more)
            {
                var results = inbox.FindItems(view);
                foreach (Item item in results)
                {
                    var props = new PropertySet(ItemSchema.MimeContent);
                    
                    // This results in a GetItem call to EWS.
                    EmailMessage email = EmailMessage.Bind(service, item.Id, props);

                    var emlFileName = $"{path}{Path.PathSeparator}email-{item.DateTimeCreated:yyyyMMddHHmmss}-{item.DateTimeReceived:yyyyMMddHHmmss}.eml";
                    // Save as .eml.
                    using (FileStream fs = new FileStream(emlFileName, FileMode.Create, FileAccess.Write))
                    {
                        fs.Write(email.MimeContent.Content, 0, email.MimeContent.Content.Length);
                    }
                }

                more = results.MoreAvailable;
                if (more)
                {
                    view.Offset += pageSize;
                }
            }
        }
    }
}
