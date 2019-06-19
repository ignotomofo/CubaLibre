using System.Collections.Generic;

namespace IntercambriarBiblio
{
    /// <summary>
    /// A collection class for exchange principles. Supports collection initialization.
    /// </summary>
    public class ExchangePrincipleCollection : List<ExchangePrinciple>
    {
        public void Add(string emailAddress)
        {
            Add(new ExchangePrinciple(emailAddress));
        }
    }
}
