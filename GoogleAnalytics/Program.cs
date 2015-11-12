using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace GoogleAnalytics
{
    class Program
    {
        static void Main(string[] args)
        {
            ConsoleColor defaultForeground = Console.ForegroundColor;

            string webUrl = "https://amtests.sharepoint.com/sites/sample/";
            string userName = "amuser@amtests.onmicrosoft.com";

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Enter your password.");
            Console.ForegroundColor = defaultForeground;
            SecureString password = GetPasswordFromConsoleInput();

            using (var context = new ClientContext(webUrl))
            {
                context.Credentials = new SharePointOnlineCredentials(userName, password);
                context.Load(context.Web, w => w.Title);
                context.ExecuteQuery();
                var baseFolder = System.IO.Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);
                string filename = System.IO.Path.Combine(baseFolder, "Files\\googletracking.js");
               // string filename = @"c:\users\aravind\documents\visual studio 2015\Projects\GoogleAnalytics\GoogleAnalytics\Files\googletracking.js";
                Console.WriteLine("Adding the Tracking javascript to the Style Library");
                var UploadStatus =UploadFile(context, "Style Library", filename);
                if (UploadStatus)
                {
                    Console.WriteLine("Adding the Scriptlinks to the specified site collection");
                    AddCustomAction(context, "ScriptLink-GoogleAnalytics", "~sitecollection/Style Library/googletracking.js", 10000);
                }

            }
        }
        private static SecureString GetPasswordFromConsoleInput()
        {
            ConsoleKeyInfo info;

            //Get the user's password as a SecureString
            SecureString securePassword = new SecureString();
            do
            {
                info = Console.ReadKey(true);
                if (info.Key != ConsoleKey.Enter)
                {
                    securePassword.AppendChar(info.KeyChar);
                }
            }
            while (info.Key != ConsoleKey.Enter);
            return securePassword;
        }
        private static Boolean UploadFile(ClientContext context, string listTitle, string fileName)
        {
            using (var fs = new FileStream(fileName, FileMode.Open))
            {
                var fi = new FileInfo(fileName);
                var list = context.Web.Lists.GetByTitle(listTitle);
                context.Load(list.RootFolder);
                context.ExecuteQuery();
                if (list != null)
                {
                    var fileUrl = String.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, fi.Name);
                    if (context.HasPendingRequest)
                        context.ExecuteQuery();
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(context, fileUrl, fs, true);
                    return true;
                }
                else
                {
                    Console.WriteLine("Library not Found");
                    return false;

                }
                
            }
        }
        private static void AddCustomAction(ClientContext context, string name, string scriptPrefixedUrl, int sequence)
        {
            var site = context.Site;
            if (!site.IsObjectPropertyInstantiated("UserCustomActions"))
            {
                context.Load(site.UserCustomActions, collection => collection.Include(ca => ca.Name));
                context.ExecuteQuery();
            }
            var action = site.UserCustomActions.FirstOrDefault(ca => string.Equals(ca.Name, name, StringComparison.InvariantCultureIgnoreCase));
            if (action == null)
            {
                action = site.UserCustomActions.Add();
                action.Location = "ScriptLink";
                action.Name = name;
                action.ScriptSrc = scriptPrefixedUrl;
                action.Sequence = sequence;
                action.Update();
                context.ExecuteQuery();
            }
            else {
                //Remove existing Site Action
                action.DeleteObject();
                context.Load(site.UserCustomActions);
                context.ExecuteQuery();
                //Add the Custom Action with the updated values
                action = site.UserCustomActions.FirstOrDefault(ca => string.Equals(ca.Name, name, StringComparison.InvariantCultureIgnoreCase));
                if (action == null)
                {
                    action = site.UserCustomActions.Add();
                    action.Location = "ScriptLink";
                    action.Name = name;
                    action.ScriptSrc = scriptPrefixedUrl;
                    action.Sequence = sequence;
                    action.Update();
                    context.ExecuteQuery();
                }
            }
        
        }
    }
}
