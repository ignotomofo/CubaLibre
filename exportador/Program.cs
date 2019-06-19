﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using IntercambriarBiblio.Mail;
using Microsoft.Exchange.WebServices.Data;

namespace exportador
{
    class Program
    {
        private static readonly string Fs = Path.DirectorySeparatorChar.ToString();
        private static readonly string Dir = AppDomain.CurrentDomain.BaseDirectory;
        private static string _export = "export";
        static void Main(string[] args)
        {
            bool isInteractive = args.Contains("-i");

            var accounts = TargetAccountSection.GetTargets();
            if (accounts == null || accounts.Count == 0)
            {
                isInteractive = true;
                Console.WriteLine("No accounts have been configure, going to interactive mode.");
            }

            if (isInteractive)
            {
                accounts = GetAccountInteractively();
            }

            var d = ConfigurationManager.AppSettings["exportFolder"];
            if (!string.IsNullOrWhiteSpace(d))
            {
                _export = d;
            }

            foreach (var account in accounts)
            {
                try
                {
                    ExportTarget(account);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    if (ex.InnerException != null)
                    {
                        Console.WriteLine(ex.InnerException.Message);
                        if (ex.InnerException.InnerException != null)
                        {
                            Console.WriteLine(ex.InnerException.InnerException.Message);
                        }
                    }
                }
            }
            Console.WriteLine("Export Complete!");
        }

        private static List<Target> GetAccountInteractively()
        {
            var userName = "";
            var password = "";
            Console.Write("Email: ");
            userName = Console.ReadLine();
            Console.Write("Password: ");
            password = Console.ReadLine();
            return new List<Target>
            {
                new Target
                {
                    Account = userName,
                    Password = password
                }
            };
        }

        private static void ExportTarget(Target t)
        {
            var exportPath = $"{Dir}{Fs}{_export}{Fs}{t.Account}";
            if (Directory.Exists(exportPath))
            {
                Directory.Delete(exportPath, true);
            }
            Directory.CreateDirectory(exportPath);
            var uri = System.Configuration.ConfigurationManager.AppSettings["uri"];
            var s = (string.IsNullOrWhiteSpace(uri))
                ? new MailService(t.Account, t.Password)
                : new MailService(t.Account, t.Password, new Uri(uri));
            Console.WriteLine($"Exporting to {t.Account} Inbox");
            s.ExportMail(exportPath + $"{Fs}Inbox", WellKnownFolderName.Inbox);
            Console.WriteLine($"Exporting to {t.Account} SentItems");
            s.ExportMail(exportPath + $"{Fs}SentItems", WellKnownFolderName.SentItems);
            Console.WriteLine($"Complete.");

        }
    }
}