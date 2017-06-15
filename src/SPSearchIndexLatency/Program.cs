using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;

namespace SPSearchIndexLatency
{
    class Program
    {
        static ClientContext SetupContext(Uri siteCollection, string username, string password)
        {
            if (string.IsNullOrEmpty(username) && string.IsNullOrEmpty(password))
            {
                return new ClientContext(siteCollection);
            }
            else
            {
                var securePassword = new SecureString();
                password.ToCharArray().ToList().ForEach(securePassword.AppendChar);
                var credentials = new SharePointOnlineCredentials(username, securePassword);
                return new ClientContext(siteCollection) { Credentials = credentials };
            }
        }

        static void Ping(List pings)
        {
            var ping = pings.AddItem(new ListItemCreationInformation());
            ping["Title"] = DateTime.Now.ToUniversalTime();
            ping.Update();
            pings.Context.ExecuteQuery();
        }

        static void QuerySearchEngine(ClientContext ctx, string kql, Action<IDictionary<string, object>> processRow, int startRow = 0)
        {
            const int BatchSize = 500;
            var executor = new SearchExecutor(ctx);
            var results = executor.ExecuteQuery(
                new KeywordQuery(ctx)
                {
                    QueryText = kql,
                    StartRow = startRow,
                    RowLimit = BatchSize
                });
            ctx.ExecuteQuery();

            var rows = results.Value[0];
            rows.ResultRows.ToList().ForEach(processRow);

            if (rows.RowCount > 0)
            {
                QuerySearchEngine(ctx, kql, processRow, startRow + BatchSize);
            }
        }

        static void UpdateResults(List<DateTime> receivedPings, int run)
        {
            var replies = receivedPings.Count();
            var lastReceived = replies == 0 ? "N/A\t\t" : receivedPings.Max().ToString(@"MM\/dd\/yyyy HH:mm:ss");
            var now = DateTime.Now.ToUniversalTime();
            var latency = replies == 0 ? "N/A" : (now - receivedPings.Max()).ToString(@"hh\:mm\:ss");

            if (run == 0)
            {
                Console.WriteLine("Run\tReplies\t\tCurrent time\t\tMost recent reply\tLatency");
            }

            Console.WriteLine(
                "{0}\t{1}\t\t{2}\t{3}\t{4}",
                run + 1,
                replies,
                now.ToString(@"MM\/dd\/yyyy HH:mm:ss"),
                lastReceived,
                latency);
        }

        static void Main(string[] args)
        {
            if (args.Length < 3 && args.Length != 5)
            {
                Console.WriteLine("SPSearchIndexLiveliness.Console.exe webUrl listTitle pingInterval [username] [password]");
                Console.WriteLine("SPSearchIndexLiveliness.Console.exe https://bugfree.sharepoint.com/sites/liveliness Pings 30 rh@bugfree.onmicrosoft.com password");
                return;
            }

            var webUrl = new Uri(args[0]);
            var listTitle = args[1];
            var pingInterval = int.Parse(args[2]) * 1000;

            string username = null;
            string password = null;
            if (args.Length == 5)
            {
                username = args[3];
                password = args[4];
            }

            var receivedPings = new List<DateTime>();
            Action<IDictionary<string, object>> processRow = r => receivedPings.Add((DateTime)r["Write"]);

            using (var ctx = SetupContext(webUrl, username, password))
            {
                var library = ctx.Web.Lists.GetByTitle(listTitle);
                ctx.Load(library);
                ctx.ExecuteQuery();

                var runs = 0;
                while (true)
                {
                    var watch = new Stopwatch();
                    watch.Start();

                    try
                    {
                        Ping(library);
                        QuerySearchEngine(ctx, "contentclass:STS_ListItem ListID:" + library.Id, processRow);
                        UpdateResults(receivedPings, runs);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"{e.Message}{Environment.NewLine}{e.StackTrace}");
                    }
                    finally
                    {
                        runs++;
                        receivedPings.Clear();
                    }

                    var waitTime = pingInterval - (int)watch.ElapsedMilliseconds;
                    System.Threading.Thread.Sleep(waitTime < 0 ? 0 : waitTime);
                }
            }
        }
    }
}