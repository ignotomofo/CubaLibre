
namespace IntercambriarBiblio
{
    /// <summary>
    /// Represents an exchange account. This is to be used for things
    /// like sending email or creating calendar events. 
    /// </summary>
    public class ExchangePrinciple
    {
        /// <summary>
        /// Id to be used with ExchangeServices.
        /// </summary>
        public string Id { get; }

        /// <summary>
        /// Constructs the ExchangePrinciple from email address.    
        /// </summary>
        /// <param name="emailAddress"></param>
        public ExchangePrinciple(string emailAddress)
        {
            Id = emailAddress;
        }
    }
}
