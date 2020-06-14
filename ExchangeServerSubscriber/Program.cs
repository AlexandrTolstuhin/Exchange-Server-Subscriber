using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using static System.Console;

namespace ExchangeServerSubscriber
{
    internal sealed class Program
    {
        private static readonly IWatermarkStorage WatermarkStorage = new WatermarkStorage();

        private static void Main()
        {
            var service = new ExchangeService
            {
                Url = new Uri(ConfigurationManager.AppSettings["exchangeServiceUri"]),
                Credentials = new WebCredentials(
                    ConfigurationManager.AppSettings["username"], 
                    ConfigurationManager.AppSettings["password"])
            };

            SetStreamingNotifications(service);

            ConsoleKeyInfo key;
            do
            {
                key = ReadKey(true);
            } while (key.Key != ConsoleKey.Escape);
        }

        static void SetStreamingNotifications(ExchangeService service)
        {
            var streamingSubscription = service.SubscribeToStreamingNotifications(
                new FolderId[] { WellKnownFolderName.Inbox },
                EventType.NewMail);

            var connection = new StreamingSubscriptionConnection(service, 30);

            connection.AddSubscription(streamingSubscription);

            connection.OnNotificationEvent += OnEvent;
            connection.OnSubscriptionError += OnError;
            connection.OnDisconnect += OnDisconnect;
            connection.Open();

            SynchronizeChanges(service);
        }

        private static void OnDisconnect(object sender, SubscriptionErrorEventArgs args)
        {
            var connection = (StreamingSubscriptionConnection)sender;

            WriteLine($"Date: {DateTime.Now:HH:mm:ss}");
            WriteLine("The connection to the subscription is disconnected.");
            WriteLine("Reconnect the subscription.");

            connection.Open();
        }

        static void OnEvent(object sender, NotificationEventArgs args)
            => SynchronizeChanges(args.Subscription.Service);


        static void OnError(object sender, SubscriptionErrorEventArgs args)
            => WriteLine($"{args.Exception.Message}");

        public static void SynchronizeChanges(ExchangeService service)
        {
            bool moreChangesAvailable;

            var property = new PropertySet(BasePropertySet.FirstClassProperties)
            {
                RequestedBodyType = BodyType.Text
            };

            WriteLine("Synchronizing changes...");

            var messages = new List<EmailMessage>();
            var syncState = WatermarkStorage.Load();

            do
            {
                var changes = service.SyncFolderItems(
                    WellKnownFolderName.Inbox,
                    PropertySet.FirstClassProperties,
                    null, 512,
                    SyncFolderItemsScope.NormalItems,
                    syncState);

                syncState = changes.SyncState;

                foreach (var itemChange in changes.Where(_ => _.ChangeType == ChangeType.Create))
                {
                    if (itemChange.Item.DateTimeCreated < DateTime.Today) break;

                    try
                    {
                        var message = EmailMessage.Bind(service, itemChange.ItemId, property);

                        messages.Add(message);
                    }
                    catch (Exception e)
                    {
                        WriteLine(e.Message);
                    }
                }

                moreChangesAvailable = changes.MoreChangesAvailable;
            } while (moreChangesAvailable);

            foreach (var message in messages.OrderBy(_ => _.DateTimeCreated))
            {
                WriteLine();
                WriteLine("New mail:");
                WriteLine($"Date: {message.DateTimeCreated}");
                WriteLine($"From: {message.From}");
                WriteLine($"Subject: {message.Subject}");
                WriteLine($"HasAttachments: {message.HasAttachments}");
                WriteLine($"Attachments.Count: {message.Attachments.Count}");
                WriteLine();
            }

            WatermarkStorage.Save(syncState);
            WriteLine("Done.");
        }
    }
}