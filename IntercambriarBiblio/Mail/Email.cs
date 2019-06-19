namespace IntercambriarBiblio.Mail
{
    public enum EmailBodyType
    {
        Text,
        Html
    };

    public class Email
    {

        public ExchangePrincipleCollection ToRecipients { get; }
        public ExchangePrincipleCollection CcRecipients { get; }
        public ExchangePrincipleCollection BccRecipients { get; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public EmailBodyType BodyType { get; set; }

        public Email()
        {
            ToRecipients = new ExchangePrincipleCollection();
            CcRecipients = new ExchangePrincipleCollection();
            BccRecipients = new ExchangePrincipleCollection();
            BodyType = EmailBodyType.Text;
        }           
    }
}
