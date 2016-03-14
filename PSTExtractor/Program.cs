using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Redemption;

namespace PSTExtractor
{
    static class Program
    {
        static void Main(string[] args)
        {
            var dataFolder = Environment.CurrentDirectory + @"\..\..\data";
            var pstFilePath = Path.Combine(dataFolder, "Newsletter.pst");
            var outputFilePath = Path.Combine(dataFolder, "recipients-failed.csv");


            if (!File.Exists(pstFilePath))
            {
                Console.WriteLine("File not found at " + pstFilePath);
                Console.ReadLine();
                return;
            }

           
            Console.WriteLine("Input file: {0}", pstFilePath);

            //var options = RegexOptions.Multiline & RegexOptions.IgnoreCase;
            //var temporarily = @"try again|temporarily|quota|exceeded";
            //var permanent = @"refused|rejected|unknown|no such user|denied";
            
            try
            {
                var reportItems = GetReportItems(pstFilePath);
                var failedRecipients = reportItems.Select(x => x.Recipients);
                
                var failedMails = GetAddressList(failedRecipients);
                Console.WriteLine("Found {0} matching emails for failed recipients.", failedMails.Count());

                ExportToCsv(failedMails, outputFilePath);
            }
            catch (SystemException ex)
            {
                Console.WriteLine(ex.Message);
            }

            Console.ReadLine();
        }

        private static List<string> GetAddressList(IEnumerable<RDORecipients> recipients)
        {
            var addressList = new List<string>();
            foreach (var recipient in recipients)
            {
                var count = recipient.Count;
                for (var i = 1; i <= count; i++)
                {
                    addressList.Add(recipient.Item(i).Address);
                }
            }
            return addressList;
        }

        private static IEnumerable<RDOReportItem> GetReportItems(string pstFilePath)
        {
            var app = new Application();
            var rdoSession = new RDOSession { MAPIOBJECT = app.Session.MAPIOBJECT };
            
            var outlookNs = app.GetNamespace("MAPI");

            // Add PST file (Outlook Data File) to Default Profile
            outlookNs.AddStore(pstFilePath);

            var folder =(Folder) outlookNs.Folders.GetLast();

            var reportItems = folder.Items.OfType<ReportItem>()
                .Select(item => rdoSession.GetRDOObjectFromOutlookObject(item))
                .Cast<RDOReportItem>()
                .ToList();

            // Remove PST file from Default Profile
            outlookNs.RemoveStore(folder);

            return reportItems;
        }

        private static void ExportToCsv(IEnumerable<string> emails, string outputFilePath)
        {   
            const string delimiter = ";";

            if (!File.Exists(outputFilePath))
            {
                File.Create(outputFilePath).Close();
            }

            var output = emails.ToArray();
            var length = output.GetLength(0);

            var csv = new StringBuilder();  
	            for (var index = 0; index < length; index++)
                    csv.AppendLine(string.Join(delimiter, output[index]));

            File.WriteAllText(outputFilePath, csv.ToString());
            Console.WriteLine("Saved to file: {0}", outputFilePath);
        }

    }
}
