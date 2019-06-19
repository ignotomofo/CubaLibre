using System;
using System.Collections.Generic;
using System.Configuration;
using System.Xml;

namespace exportador
{
    public class TargetAccountSection : IConfigurationSectionHandler
    {
        public object Create(object parent, object configContext, XmlNode section)
        {
            var listOfTargets = new List<Target>();

            foreach (XmlNode childNode in section.ChildNodes)
            {
                var t = new Target();
                if (childNode.Attributes != null)
                {
                    foreach (XmlAttribute attribute in childNode.Attributes)
                    {
                        if (attribute.Name == "account")
                        {
                            t.Account = attribute.Value;
                        }
                        else if (attribute.Name == "password")
                        {
                            t.Password = attribute.Value;
                        }


                    }
                    listOfTargets.Add(t);
                }
            }
            return listOfTargets;
        }

        public static List<Target> GetTargets()
        {
            var l = new List<Target>();
            try
            {
                if (ConfigurationManager.GetSection("targetAccounts") is List<Target> accounts)
                {
                    l.AddRange(accounts);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Configuration Error: {ex}");
            }
            return l;
        }
    }

    public class Target
    {
        public string Account { get; set; }
        public string Password { get; set; }

    }
}
