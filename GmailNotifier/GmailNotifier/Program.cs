using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using System.Drawing;
using System.Windows.Forms;
using Thread = Google.Apis.Gmail.v1.Data.Thread;
using System.Collections.Generic;

namespace GmailTest
{
    class Program
    {
        private static NotifyIcon TrayIcon;
        private static Dictionary<string, Google.Apis.Gmail.v1.Data.Message> UnreadMessagesNotifications;
        private const string ClientId = "INSERT CLIENT ID HERE";
        private const string ClientSecret = "INSERT CLIENT SECRET HERE";

        static void Main(string[] args)
        {
            TrayIcon = new NotifyIcon();
            TrayIcon.BalloonTipIcon = ToolTipIcon.None;
            TrayIcon.Icon = GmailNotifier.Properties.Resources.ReadGmail32;
            TrayIcon.Visible = true;

            UnreadMessagesNotifications = new Dictionary<string, Google.Apis.Gmail.v1.Data.Message>();

            while (true)
            {
                try
                {
                    new Program().CheckGmail().Wait();
                }
                catch (AggregateException ex)
                {
                    foreach (var e in ex.InnerExceptions)
                    {
                        Console.WriteLine("Exception: " + e.Message);
                    }
                }
            }
        }
        private async Task CheckGmail()
        {

            UserCredential credential;
            ClientSecrets secrets = new ClientSecrets();
            secrets.ClientId = ClientId;
            secrets.ClientSecret = ClientSecret;
            credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(secrets,
                new[] { GmailService.Scope.GmailReadonly },
                "user", CancellationToken.None);

            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Gmail Notifier",
            });

            while (true)
            {
                try
                {
                    var response = service.Users.Labels.Get("me", "INBOX").Execute();
                    if (response.MessagesUnread > 0)
                    {
                        TrayIcon.Icon = GmailNotifier.Properties.Resources.UnreadGmail32;
                        if (response.MessagesUnread == 1)
                        {
                            TrayIcon.Text = "Du har ett oläst meddelande!";
                        }
                        else
                        {
                            TrayIcon.Text = String.Format("Du har {0} olästa meddelanden!", response.MessagesUnread);
                        }
                        TrayIcon.Visible = true;
                        
                        var unreadMessages = service.Users.Messages.List("me");
                        unreadMessages.LabelIds = new[] { "INBOX", "UNREAD" };
                        var inbox = unreadMessages.Execute();

                        foreach (Google.Apis.Gmail.v1.Data.Message message in inbox.Messages)
                        {
                            if (!UnreadMessagesNotifications.ContainsKey(message.Id))
                            {
                                var mess = service.Users.Messages.Get("me", message.Id).Execute();
                                var subject = mess.Payload.Headers.Where(x => x.Name == "Subject").FirstOrDefault();
                                TrayIcon.ShowBalloonTip(3000, subject.Value, mess.Snippet, ToolTipIcon.Info);
                                TrayIcon.Icon = GmailNotifier.Properties.Resources.NewGmail32;
                                UnreadMessagesNotifications.Add(message.Id, message);
                                System.Threading.Thread.Sleep(5000);
                            }
                        }
                        TrayIcon.Icon = GmailNotifier.Properties.Resources.UnreadGmail32;
                    }
                    else
                    {
                        TrayIcon.Icon = GmailNotifier.Properties.Resources.ReadGmail32;
                        TrayIcon.Text = "Inga olästa meddelande!";
                    }
                    System.Threading.Thread.Sleep(500);
                }
                catch (Exception e)
                {
                    Console.WriteLine("An error occurred: " + e.Message);
                }
            }
        }

    }
}