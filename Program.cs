/*
 * By David Barrett, Microsoft Ltd. 2021. Use at your own risk.  No warranties are given.
 * 
 * DISCLAIMER:
 * THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
 * MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
 * A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
 * MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
 * BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
 * SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
 * OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.
 * */

using System;
using Microsoft.Identity.Client;
using System.Threading.Tasks;
using System.Net.Sockets;
using System.Text;
using System.Net.Security;

namespace IMAPOAuthSample
{
    class Program
    {

        private static TcpClient _imapClient = null;
        private static SslStream _sslStream = null;

        static void Main(string[] args)
        {
            if (args.Length<2)
            {
                Console.WriteLine($"Syntax: {System.Reflection.Assembly.GetExecutingAssembly().GetName()} <TenantId> <ApplicationId>");
                return;
            }
            //TestIMAP(args[0], args[1]);
            var program = new Program();
            var task = TestIMAP(args[1], args[0]);
            //task.RunSynchronously();
            task.Wait();
        }


        static async Task TestIMAP(string ClientId, string TenantId)
        {

            // Configure the MSAL client to get tokens
            var pcaOptions = new PublicClientApplicationOptions
            {
                ClientId = ClientId,
                TenantId = TenantId
            };

            Console.WriteLine("Building application");
            var pca = PublicClientApplicationBuilder
                .CreateWithApplicationOptions(pcaOptions)
                .WithRedirectUri("http://localhost")
                .Build();

            var imapScope = new string[] { "https://outlook.office.com/IMAP.AccessAsUser.All" };

            try
            {
                // Make the interactive token request
                Console.WriteLine("Requesting access token (user must log-in via browser)");
                var authResult = await pca.AcquireTokenInteractive(imapScope).ExecuteAsync();
                if (String.IsNullOrEmpty(authResult.AccessToken))
                {
                    Console.WriteLine("No token received");
                    return;
                }
                Console.WriteLine($"Token received for {authResult.Account.Username}");


                RetrieveMessages(authResult);
            }
            catch (MsalException ex)
            {
                Console.WriteLine($"Error acquiring access token: {ex}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex}");
            }
            Console.WriteLine("Finished");

        }

        static string ReadSSLStream()
        {
            int bytes = -1;
            byte[] buffer = new byte[2048];
            bytes = _sslStream.Read(buffer, 0, buffer.Length);
            string response = Encoding.ASCII.GetString(buffer, 0, bytes);
            Console.WriteLine(response);
            return response;
        }

        static void WriteSSLStream(string Data)
        {
            _sslStream.Write(Encoding.ASCII.GetBytes($"{Data}{Environment.NewLine}"));
            _sslStream.Flush();
            Console.WriteLine(Data);
        }

        static void RetrieveMessages(AuthenticationResult authResult)
        {
            try
            {
                _imapClient = new TcpClient("outlook.office365.com", 993);
                _sslStream = new SslStream(_imapClient.GetStream());
                _sslStream.AuthenticateAsClient("outlook.office365.com");

                ReadSSLStream();

                //Send the users login details
                WriteSSLStream($"$ CAPABILITY");
                ReadSSLStream();

                //Send the users login details
                WriteSSLStream($"$ AUTHENTICATE XOAUTH2 {XOauth2(authResult)}");
                string response = ReadSSLStream();
                if (response.StartsWith("$ NO AUTHENTICATE"))
                    Console.WriteLine("Authentication failed");
                else
                {
                    // Retrieve inbox unread messages
                    WriteSSLStream("$ STATUS INBOX (unseen)");
                    ReadSSLStream();

                    // Log out
                    WriteSSLStream($"$ LOGOUT");
                    ReadSSLStream();
                }


                // Tidy up
                Console.WriteLine("Closing connection");
                _sslStream.Close();
                _sslStream.Close();
            }
            catch (SocketException ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        static string XOauth2(AuthenticationResult authResult)
        {
            string ctrlA = $"{(char)1}";
            string login = $"user={authResult.Account.Username}{ctrlA}auth=Bearer {authResult.AccessToken}{ctrlA}{ctrlA}";
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(login);
            return Convert.ToBase64String(plainTextBytes);
        }
    }
}
