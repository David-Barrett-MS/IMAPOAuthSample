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
        private static string _imapEndpoint = "outlook.office365.com";

        static void Main(string[] args)
        {
            if (args.Length<2)
            {
                Console.WriteLine($"Syntax: {System.Reflection.Assembly.GetExecutingAssembly().GetName()} <TenantId> <ApplicationId> <SecretKey> <Mailbox>");
                return;
            }

            Task task = null;
            if (args.Length > 2 && !String.IsNullOrEmpty(args[2]) && !String.IsNullOrEmpty(args[3]))
            {
                // Application auth
                task = TestIMAP(args[1], args[0], args[2], args[3]);
            }
            else
                task = TestIMAP(args[1], args[0]);

            task.Wait();
        }


        static async Task TestIMAP(string ClientId, string TenantId, string SecretKey=null, string mailbox=null)
        {
            var imapScope = new string[] { $"https://{_imapEndpoint}/IMAP.AccessAsUser.All" };

            Console.WriteLine("Building OAuth application");
            if (String.IsNullOrEmpty(SecretKey))
            {
                // Configure the MSAL client to get tokens
                var pcaOptions = new PublicClientApplicationOptions
                {
                    ClientId = ClientId,
                    TenantId = TenantId
                };

                // Interactive sign-in
                var pca = PublicClientApplicationBuilder
                    .CreateWithApplicationOptions(pcaOptions)
                    .WithRedirectUri("http://localhost")
                    .Build();

                try
                {
                    // Make the interactive token request
                    Console.WriteLine("Requesting access token (user must log-in via browser)");
                    var authResult = await pca.AcquireTokenInteractive(imapScope).ExecuteAsync();
                    if (String.IsNullOrEmpty(authResult.AccessToken))
                        Console.WriteLine("No token received");
                    else
                    {
                        Console.WriteLine($"Token received for {authResult.Account.Username}");

                        // Use the token to connect to IMAP service
                        RetrieveMessages(authResult);
                    }
                }
                catch (MsalException ex)
                {
                    Console.WriteLine($"Error acquiring access token: {ex}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex}");
                }
            }
            else
            {
                // Application credentials (no user log-in required, we use secret key to obtain token)
                var cca = ConfidentialClientApplicationBuilder.Create(ClientId)
                    .WithAuthority(AzureCloudInstance.AzurePublic, TenantId)
                    .WithClientSecret(SecretKey)
                    .Build();
                imapScope = new string[] { $"https://{_imapEndpoint}/.default" };
                try
                {
                    // Acquire the token
                    Console.WriteLine("Requesting access token");
                    var authResult = await cca.AcquireTokenForClient(imapScope).ExecuteAsync();
                    Console.WriteLine($"Token received");

                    // Use the token to connect to IMAP service
                    RetrieveMessages(authResult, mailbox);
                }
                catch (MsalException ex)
                {
                    Console.WriteLine($"Error acquiring access token: {ex}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex}");
                }
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

        static void RetrieveMessages(AuthenticationResult authResult, string mailbox = null)
        {
            try
            {
                using (_imapClient = new TcpClient(_imapEndpoint, 993))
                {
                    using (_sslStream = new SslStream(_imapClient.GetStream()))
                    {
                        _sslStream.AuthenticateAsClient(_imapEndpoint);

                        ReadSSLStream();

                        //Send the users login details
                        WriteSSLStream($"$ CAPABILITY");
                        ReadSSLStream();

                        //Send the users login details
                        WriteSSLStream($"$ AUTHENTICATE XOAUTH2 {XOauth2(authResult, mailbox)}");
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
                    }
                }
            }
            catch (SocketException ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        static string XOauth2(AuthenticationResult authResult, string mailbox = null)
        {
            // Create the log-in code, which is a base 64 encoded combination of mailbox (user) and auth token

            char ctrlA = (char)1;
            if (String.IsNullOrEmpty(mailbox))
                mailbox = authResult.Account.Username;
            Console.WriteLine($"Authenticating for access to mailbox {mailbox}");
            string login = $"user={mailbox}{ctrlA}auth=Bearer {authResult.AccessToken}{ctrlA}{ctrlA}";
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(login);
            return Convert.ToBase64String(plainTextBytes);
        }
    }
}
