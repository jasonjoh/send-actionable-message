using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net.Http.Headers;
using System.Threading.Tasks;

using Microsoft.Identity.Client;
using Microsoft.Graph;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using System.Globalization;

namespace send_actionable_message
{
    class Program
    {
        static PublicClientApplication authClient = null;
        static string[] scopes =
        {
            "User.Read", // Scope needed to read /Me from Graph (to get email address)
            "Mail.Send"  // Scope needed to send mail as the user
        };

        static void Main(string[] args)
        {
            SendMessageAsync(args).Wait();
            Console.WriteLine("Hit any key to exit...");
            Console.ReadKey();
        }

        static async Task SendMessageAsync(string[] args)
        {
            // Setup MSAL client
            authClient = new PublicClientApplication(ConfigurationManager.AppSettings.Get("applicationId"));

            try
            {
                // Get the access token
                var result = await authClient.AcquireTokenAsync(scopes);

                // Initialize Graph client with delegate auth provider
                // that just returns the token we already retrieved
                var graphClient = new GraphServiceClient(
                    new DelegateAuthenticationProvider(
                        (requestMessage) =>
                        {
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
                            return Task.FromResult(0);
                        }));

                // Create a recipient
                var me = await graphClient.Me.Request().GetAsync();
                var toRecip = new Recipient()
                {
                    EmailAddress = new EmailAddress() {
                        // If recipient provided as an argument, use that
                        // If not, use the logged in user
                        Address = IsValidEmail(args) ? args[0] : me.Mail
                    }
                };

                // Create the message
                var actionableMessage = new Message()
                {
                    Subject = "Actionable message sent from code",
                    ToRecipients = new List<Recipient>() { toRecip },
                    Body = new ItemBody()
                    {
                        ContentType = BodyType.Html,
                        Content = LoadActionableMessageBody()
                    },
                    Attachments = new MessageAttachmentsCollectionPage()
                };

                // Create an attachment for the activity image
                var actionImage = new FileAttachment()
                {
                    ODataType = "#microsoft.graph.fileAttachment",
                    Name = "activity_image", // IMPORTANT: Name must match ContentId
                    IsInline = true,
                    ContentId = "activity_image",
                    ContentType = "image/jpg",
                    ContentBytes = System.IO.File.ReadAllBytes(@".\ActivityImage.jpg")
                };

                actionableMessage.Attachments.Add(actionImage);

                // Send the message
                await graphClient.Me.SendMail(actionableMessage, true).Request().PostAsync();

                Output.WriteLine(Output.Success, "Message sent");
            }
            catch (MsalException ex)
            {
                Output.WriteLine(Output.Error, "An exception occurred while acquiring an access token.");
                Output.WriteLine(Output.Error, "  Code: {0}; Message: {1}", ex.ErrorCode, ex.Message);
            }
            catch (Microsoft.Graph.ServiceException graphEx)
            {
                Output.WriteLine(Output.Error, "An exception occurred while making a Graph request.");
                Output.WriteLine(Output.Error, "  Code: {0}; Message: {1}", graphEx.Error.Code, graphEx.Message);
            }
        }

        // Copied from https://docs.microsoft.com/dotnet/standard/base-types/how-to-verify-that-strings-are-in-valid-email-format
        static bool IsValidEmail(string[] args)
        {
            if (args.Length <= 0)
            {
                return false;
            }

            var email = args[0];
            if (string.IsNullOrEmpty(email))
            {
                return false;
            }

            // Handle any Unicode domains
            try
            {
                email = Regex.Replace(email, @"(@)(.+)$", DomainMapper,
                    RegexOptions.None, TimeSpan.FromMilliseconds(200));
            }
            catch (RegexMatchTimeoutException)
            {
                return false;
            }
            catch (ArgumentException)
            {
                return false;
            }

            try
            {
                return Regex.IsMatch(email,
                    @"^(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                    @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-0-9a-z]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9]))$",
                    RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250));
            }
            catch (RegexMatchTimeoutException)
            {
                return false;
            }
        }

        static string DomainMapper(Match match)
        {
            var idn = new IdnMapping();

            string domainName = match.Groups[2].Value;
            domainName = idn.GetAscii(domainName);

            return $"{match.Groups[1].Value}{domainName}";
        }

        static string LoadActionableMessageBody()
        {
            // Load the card JSON
            var cardJson = JObject.Parse(System.IO.File.ReadAllText(@".\Card.json"));

            // Check type
            // First, try "@type", which is the key MessageCard uses
            var cardType = cardJson.SelectToken("@type");
            if (cardType == null)
            {
                // Maybe it's Adaptive, try "type"
                cardType = cardJson.SelectToken("type");
            }

            // If we're still null, or the values are bad, bail
            if (cardType == null || (cardType.ToString() != "MessageCard" && cardType.ToString() != "AdaptiveCard"))
            {
                throw new ArgumentException("The payload in Card.json is missing a valid @type or type property.");
            }

            string scriptType = cardType.ToString() == "MessageCard" ? "application/ld+json" : "application/adaptivecard+json";

            // Insert originator if one is configured
            string originatorId = ConfigurationManager.AppSettings.Get("originatorId");
            if (!string.IsNullOrEmpty(originatorId))
            {
                // First check if there is an existing originator value
                var originator = cardJson.SelectToken("originator");

                if (originator != null)
                {
                    // Overwrite existing value
                    cardJson["originator"] = originatorId;
                }
                else
                {
                    // Add value
                    cardJson.Add(new JProperty("originator", originatorId));
                }
            }

            // Insert the JSON into the HTML
            return string.Format(System.IO.File.ReadAllText(@".\MessageBody.html"), scriptType, cardJson.ToString());
        }
    }
}
