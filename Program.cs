/*
 * By David Barrett, Microsoft Ltd. 2022-2023. Use at your own risk.  No warranties are given.
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
        /// <summary>
        /// TcpClient for the IMAP connection
        /// </summary>
        private static TcpClient _imapClient = null;
        /// <summary>
        /// SSLStream for communicating with IMAP server
        /// </summary>
        private static SslStream _sslStream = null;
        /// <summary>
        /// IMAP server address
        /// </summary>
        private static string _imapEndpoint = "outlook.office365.com";
        /// <summary>
        /// The redirect URL used for app auth
        /// </summary>
        private static string _redirectUrl = "http://localhost";

        /// <summary>
        /// Main entry point
        /// </summary>
        /// <param name="args">Command line arguments</param>
        static void Main(string[] args)
        {
            if (args.Length<2 || args.Length>4)
            {
                Console.WriteLine("Syntax:");
                Console.WriteLine("");
                Console.WriteLine("App auth:");
                Console.WriteLine($"{System.Reflection.Assembly.GetExecutingAssembly().GetName()} <TenantId> <ApplicationId> <SecretKey> <Mailbox>");
                Console.WriteLine("");
                Console.WriteLine("Delegated (interactive) auth:");
                Console.WriteLine($"{System.Reflection.Assembly.GetExecutingAssembly().GetName()} <TenantId> <ApplicationId> (<RedirectUrl>)");
                Console.WriteLine($"If <RedirectUrl> is not specified, it defaults to {_redirectUrl}");
                return;
            }

            Task task = null;
            if (args.Length > 3 && !String.IsNullOrEmpty(args[2]) && !String.IsNullOrEmpty(args[3]))
            {
                // Application auth
                task = TestIMAP(args[1], args[0], args[2], args[3]);
            }
            else
            {
                /// Interactive (delegated) auth
                if (args.Length==3 && !String.IsNullOrEmpty(args[2]))
                    _redirectUrl = args[2];
                Console.WriteLine($"Using redirect URL: {_redirectUrl}");
                task = TestIMAP(args[1], args[0]);
            }

            task.Wait();
        }


        /// <summary>
        /// Test the given provided IMAP details (attempt to obtain token and access mailbox)
        /// </summary>
        /// <param name="ClientId">Application/Client Id</param>
        /// <param name="TenantId">Tenant Id</param>
        /// <param name="SecretKey">Required for app auth</param>
        /// <param name="mailbox">Required for app auth</param>
        /// <returns>Task (no data returned, task runs until completion)</returns>
        static async Task TestIMAP(string ClientId, string TenantId, string SecretKey=null, string mailbox=null)
        {
            var imapScope = new string[] { $"https://{_imapEndpoint}/IMAP.AccessAsUser.All" };

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
                    .WithRedirectUri(_redirectUrl)
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
                    Console.WriteLine("Requesting access token (client credentials - no user interaction required)");
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

        /// <summary>
        /// Read the server response from the active SSL stream
        /// </summary>
        /// <returns>Server response as a string</returns>
        static string ReadSSLStream()
        {
            int bytes = -1;
            byte[] buffer = new byte[2048];
            bytes = _sslStream.Read(buffer, 0, buffer.Length);
            string response = Encoding.ASCII.GetString(buffer, 0, bytes);
            Console.WriteLine(response);
            return response;
        }

        /// <summary>
        /// Write the supplied data to the active SSL stream
        /// </summary>
        /// <param name="Data">The string data to be written</param>
        static void WriteSSLStream(string Data)
        {
            _sslStream.Write(Encoding.ASCII.GetBytes($"{Data}{Environment.NewLine}"));
            _sslStream.Flush();
            Console.WriteLine(Data);
        }

        /// <summary>
        /// Connect to a mailbox using the provided authentication information and show the count of unread messages
        /// </summary>
        /// <param name="authResult">Valid OAuth credentials to access the mailbox</param>
        /// <param name="mailbox">Mailbox to be accessed (required for app only flow)</param>
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

        /// <summary>
        /// Calculate and return the log-in code, which is a base 64 encoded combination of mailbox (user) and auth token
        /// </summary>
        /// <param name="authResult">Valid OAuth token</param>
        /// <param name="mailbox">If missing, mailbox will be read from the token</param>
        /// <returns>IMAP log-in code</returns>
        static string XOauth2(AuthenticationResult authResult, string mailbox = null)
        {
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
