using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace elastic_outlook_indexer
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var app = new Application();
                var ns = app.GetNamespace("MAPI");
                var folders = ns.Folders;

                foreach (var item in folders)
                {
                    var folder = item as Folder;
                    if (folder != null)
                    {
                        EnumFolder(folder, IndexFolder);
                    }
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("ERROR: {0}", ex);
            }
        }

        // elastic search server's base URI
        static Uri baseUri = new Uri("http://localhost:9234/exchange/mail/");
        static HttpClient httpClient = new HttpClient();

        // type mapping
        static Dictionary<Type, Type> typeMapping = new Dictionary<Type, Type>
        {
            {typeof(Mail), typeof(MailItem)},
            {typeof(AddressEntry), typeof(Microsoft.Office.Interop.Outlook.AddressEntry)},
            {typeof(Recipient), typeof(Microsoft.Office.Interop.Outlook.Recipient)}
        };

        // types to skip mapping because we'll do it manually
        static List<Type> typesToSkip = new List<Type>
        {
            typeof(IList<Recipient>)
        };

        static void IndexFolder(Folder folder)
        {
            foreach (var item in folder.Items)
            {
                var fromMail = item as MailItem;
                if (fromMail != null)
                {
                    var toMail = new Mail();

                    // assign all simple properties
                    Utils.AssignObject<MailItem, Mail>(
                        fromMail, toMail, typeMapping, typesToSkip);

                    // assign the collections
                    toMail.Recipients = new List<Recipient>();
                    AssignRecipients(fromMail.Recipients, toMail.Recipients);
                    toMail.ReplyRecipients = new List<Recipient>();
                    AssignRecipients(fromMail.ReplyRecipients, toMail.ReplyRecipients);

                    // convert this to json and post to elastic search
                    var json = JsonConvert.SerializeObject(toMail);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    Console.Write(toMail.Subject + "\r");
                    httpClient.PutAsync(new Uri(baseUri, toMail.EntryId), content).Wait();
                }
            }
        }

        static void AssignRecipients(Recipients source, IList<Recipient> target)
        {
            foreach (Microsoft.Office.Interop.Outlook.Recipient rcpt in source)
            {
                var recipient = new Recipient();
                Utils.AssignObject<Microsoft.Office.Interop.Outlook.Recipient, Recipient>(
                    rcpt, recipient, typeMapping, typesToSkip);
                target.Add(recipient);
            }
        }

        static void EnumFolder(Folder folder, System.Action<Folder> callback)
        {
            // recursively inspecting public folders seems to cause outlook
            // to freeze up; so we don't do it
            bool isPublicFolder = folder.Name.Contains("Public Folders");
            if (isPublicFolder)
            {
                return;
            }

            // process the folder
            callback(folder);

            foreach (var item in folder.Folders)
            {
                var f = item as Folder;
                if (f != null)
                {
                    EnumFolder(f, callback);
                }
            }
        }
    }
}
