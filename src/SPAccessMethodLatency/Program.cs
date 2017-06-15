using Microsoft.SharePoint.Client;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Xml.Linq;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Net.Http;
using System.Net.Http.Headers;
using SPFile = Microsoft.SharePoint.Client.File;

namespace SPAccessMethodLatency
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

        static AuthenticationResult GetOAuthAccessToken(Uri resource, string username, string password)
        {
            const string oauth2 = "https://login.microsoftonline.com/common/oauth2/authorize";

            // Use ClientId of Excel Power Query add-in/Microsoft Power BI as per 
            // https://support.microsoft.com/en-us/kb/3133137. The alternative is
            // to explicitly register this app with Azure AD before running it.
            var clientId = new Guid("a672d62c-fc7b-4e81-a576-e60dc46e951d");
            var authenticationContext = new AuthenticationContext(oauth2);

            var authenticationResult =
              authenticationContext.AcquireTokenAsync(
                resource.ToString(), clientId.ToString(), new UserPasswordCredential(username, password)).Result;
            return authenticationResult;
        }

        static long Write(Uri destination, string username, string password)
        {
            var original = DateTime.UtcNow.Ticks;
            using (var ctx = SetupContext(destination, username, password))
            {
                var bytes = Encoding.ASCII.GetBytes(original.ToString());
                using (var ms = new MemoryStream(bytes))
                {
                    SPFile.SaveBinaryDirect(ctx, destination.LocalPath, ms, true);
                }
            }
            return original;
        }

        static long CSOMRead(Uri destination, string username, string password)
        {
            using (var ctx = SetupContext(destination, username, password))
            {
                var fileInformation = SPFile.OpenBinaryDirect(ctx, destination.LocalPath);
                using (var sr = new StreamReader(fileInformation.Stream))
                {
                    var restored = sr.ReadToEnd();
                    return long.Parse(restored);
                }
            }
        }

        static Uri GetFileByServerRelativeUrl(Uri destination)
        {
            var s = destination.Segments;
            var url2 = destination.ToString()
                .Replace(s[s.Count() - 1], "")
                .Replace(s[s.Count() - 2], "");
            return new Uri($"{url2}_api/web/GetFileByServerRelativeUrl('{destination.LocalPath}')");
        }

        static long GetTimeLastModified(XElement body)
        {
            XNamespace ns = "http://www.w3.org/2005/Atom";
            XNamespace m = "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata";
            XNamespace d = "http://schemas.microsoft.com/ado/2007/08/dataservices";

            // When we read out metadata for a file using OData, the line of interest looks like so:
            //
            //  <d:TimeLastModified m:type="Edm.DateTime">2016-11-17T14:36:46Z</d:TimeLastModified>
            //
            // Compared to actual file content this timestamp, this timestamp slightly out of order. The slight
            // difference is caused by the small amount of time it takes the server to process the write request,
            // any clock drift between the client and server, and different time resolution on the content and
            // timestamp. Taking only processing time into account, the timestamp is slightly in the future, 
            // giving a hint of server processing performance, but clock drift may pull back the timestamp to be 
            // in the past. For as long as the drift is low and constant it's nothing to worry about.
            var timeLastModified =
                body
                    .Element(ns + "content")
                    .Element(m + "properties")
                    .Element(d + "TimeLastModified");
            return DateTime.Parse(timeLastModified.Value).ToUniversalTime().Ticks;
        }

        static Tuple<long, long> OnPremRestRead(Uri destination)
        {
            // UseDefaultCredentials must be set to true or current users's Kerberos credentials aren't passed to server.
            // We cannot use HttpClient here as it doesn't support passing along Kerberos credentials.
            using (var c = new WebClient { UseDefaultCredentials = true })
            {
                var url3 = GetFileByServerRelativeUrl(destination);
                var body = c.DownloadString(url3);
                var bodyXml = XElement.Parse(body);
                var timeLastModified = GetTimeLastModified(bodyXml);
                var content = c.DownloadString(url3 + "/openbinarystream");
                return new Tuple<long, long>(timeLastModified, long.Parse(content));
            }
        }

        static Tuple<long, long> InCloudRestRead(AuthenticationResult auth, Uri destination)
        {
            using (var c = new HttpClient())
            {
                c.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", auth.AccessToken);
                var url3 = GetFileByServerRelativeUrl(destination);
                var body = c.GetStringAsync(url3).Result;
                var bodyXml = XElement.Parse(body);
                var timeLastModified = GetTimeLastModified(bodyXml);
                var content = c.GetStringAsync(url3 + "/openbinarystream").Result;
                return new Tuple<long, long>(timeLastModified, long.Parse(content));
            }
        }

        static long OnPremHttpRead(Uri destination)
        {
            // Sometimes the contents gets interval seconds behind, but catches up in the next run. That's
            // caused by the write operation not being completed with SharePoint (and BLOB storage) before 
            // the subsequent read.
            using (var c = new WebClient { UseDefaultCredentials = true })
            { 
                var content = c.DownloadString(destination);
                return long.Parse(content);
            }
        }

        static long OnPremHttpReadRandom(Uri destination)
        {
            using (var c = new WebClient { UseDefaultCredentials = true })
            { 
                var versionedUrl = new Uri($"{destination}?version={Guid.NewGuid()}");
                var content = c.DownloadString(versionedUrl);
                return long.Parse(content);
            }
        }

        static void Main(string[] args)
        {
            if (args.Length < 2 && args.Length != 4)
            {
                Console.WriteLine("SPAccessMethodLiveliness.exe destinationLibraryUrl pingInterval [username] [password]");
                Console.WriteLine("SPAccessMethodLiveliness.exe https://bugfree.sharepoint.com/sites/liveliness Pings 30 rh@bugfree.onmicrosoft.com password");
            }

            var destination = new Uri(args[0]);
            var pingInterval = int.Parse(args[1]) * 1000;

            string username = null;
            string password = null;
            if (args.Length == 4)
            {
                username = args[2];
                password = args[3];
            }

            var inCloud = destination.ToString().ToLower().Contains("sharepoint.com");
            const float ticksPerSecond = 10000000;
            var header = inCloud
                ? "Run\tWrite\tCSOM\tREST1\tREST2"
                : "Run\tWrite\tCSOM\tHTTP1\tHTTP2\tREST1\tREST2";
            Console.WriteLine(header);

            var runs = 0;
            while (true)
            {
                var watch = new Stopwatch();
                watch.Start();

                try
                {
                    var write = Write(destination, username, password);
                    var csomRead = CSOMRead(destination, username, password);

                    if (inCloud)
                    {
                        var resource = new Uri(destination.AbsoluteUri.Replace(destination.AbsolutePath, ""));
                        var authentication = GetOAuthAccessToken(resource, username, password);
                        var restRead = InCloudRestRead(authentication, destination);
                        Console.WriteLine("{0}\t{1}\t{2:0}\t{3:0.0}\t{4:0}",
                             runs,
                             new DateTime(write),
                             (csomRead - write) / ticksPerSecond,
                             (restRead.Item1 - write) / ticksPerSecond,
                             (restRead.Item2 - write) / ticksPerSecond);
                    }
                    else
                    {
                        var httpRead = OnPremHttpRead(destination);
                        var httpReadRandom = OnPremHttpReadRandom(destination);
                        var restRead = OnPremRestRead(destination);
                        Console.WriteLine("{0}\t{1}\t{2:0}\t{3:0}\t{4:0}\t{5:0.0}\t{6:0}",
                             runs,
                             new DateTime(write),
                             (csomRead - write) / ticksPerSecond,
                             (httpRead - write) / ticksPerSecond,
                             (httpReadRandom - write) / ticksPerSecond,
                             (restRead.Item1 - write) / ticksPerSecond,
                             (restRead.Item2 - write) / ticksPerSecond);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine($"{e.Message}{Environment.NewLine}{e.StackTrace}");
                }
                finally
                {
                    runs++;
                }

                var waitTime = pingInterval - (int)watch.ElapsedMilliseconds;
                System.Threading.Thread.Sleep(waitTime < 0 ? 0 : waitTime);
            }
        }
    }
}