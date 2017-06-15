using Microsoft.SharePoint.Client;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security;
using System.Threading;
using SPFile = Microsoft.SharePoint.Client.File;

namespace SPTransferSpeed
{
    class Program
    {
        const int Megabit = 1000 * 1000;
        const int Megabyte = 1024 * 1024;

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

        static double Time(Action code)
        {
            var sw = new Stopwatch();
            sw.Start();
            code();
            sw.Stop();
            return sw.ElapsedMilliseconds / 1000.0;
        }

        static void Main(string[] args)
        {
            if (args.Length < 2 && args.Length != 4)
            {
                Console.WriteLine("SPTransferSpeed.Console.exe documentItemUrl sizeInMB [username] [password]");
                Console.WriteLine("SPTransferSpeed.Console.exe https://bugfree.sharepoint.com/sites/testweb/testdoclib/testfile 1024 rh@bugfree.onmicrosoft.com password");
                return;
            }

            var destination = new Uri(args[0]);
            var sizeInMegabytes = int.Parse(args[1]);

            string username = null;
            string password = null;
            if (args.Length == 4)
            {
                username = args[2];
                password = args[3];
            }

            var contentLength = sizeInMegabytes * Megabyte;
            var randomizedContent = new byte[contentLength];
            new Random().NextBytes(randomizedContent);

            using (var ctx = SetupContext(destination, username, password))
            {
                // internally the context uses HttpWebRequest to upload and download files.
                // Thus, when uploading or downloading large files, we may experience
                //   Unhandled Exception: System.Net.WebException: The operation has timed out
                // unless we change the default timeout period of 180 seconds.
                ctx.RequestTimeout = Timeout.Infinite;

                var uploadTime = Time(() => {
                    using (var ms = new MemoryStream(randomizedContent))
                    {
                        SPFile.SaveBinaryDirect(ctx, destination.LocalPath, ms, true);
                    }
                });

                var downloadTime = Time(() => {
                    var fileInformation = SPFile.OpenBinaryDirect(ctx, destination.LocalPath);
                    using (var sr = new StreamReader(fileInformation.Stream))
                    {
                        sr.ReadToEnd();
                    }
                });

                Console.WriteLine(
                    "{0} MB uploaded in {1:0.0} seconds at {2:0.0} Mbit/s\r\n" +
                    "{0} MB downloaded in {3:0.0} seconds at {4:0.0} Mbit/s",
                    sizeInMegabytes, uploadTime, contentLength / uploadTime * 8 / Megabit,
                    downloadTime, contentLength / downloadTime * 8 / Megabit);
            }
        }
    }
}
