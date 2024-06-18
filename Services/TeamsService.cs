using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using TeamsDemo.Services.Interfaces;
using System;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace TeamsDemo.Services
{
    public class TeamsService : ITeamsService
    {
        private static GraphServiceClient _graphClient;
        private readonly string _teamId;
        private readonly string _channelId;
        private readonly string _clientId;
        private readonly string _clientSecret;
        private readonly string _tenantId;


        public TeamsService()
        {
            try
            {
                _clientId = "enter_client_id";
                _clientSecret = "enter_client_secret";
                _tenantId = "enter_tenant_id";
                _teamId = "enter_team_id";
                _channelId = "enter_channel_id";

                var scopes = new[] { "https://graph.microsoft.com/.default" };

                var options = new ClientSecretCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                };

                var clientSecretCredential = new ClientSecretCredential(_tenantId, _clientId, _clientSecret, options);
                _graphClient = new GraphServiceClient(clientSecretCredential, scopes);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error initializing TeamsService: {ex.Message}");
            }
        }

        public async Task<string> GetAccessTokenAsync(string clientId, string clientSecret, string tenantId, string[] scopes)
        {
            IConfidentialClientApplication app = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithClientSecret(clientSecret)
                .WithAuthority(new Uri("https://login.microsoftonline.com/" + tenantId))
                .Build();

            var result = await app.AcquireTokenForClient(scopes)
                .ExecuteAsync();

            return result.AccessToken;
        }

        public async Task SendMessageAsync(string message)
        {
            try
            {
                // Creăm un mesaj pentru a trimite în canal
                var chatMessage = new ChatMessage
                {
                    Body = new ItemBody
                    {
                        Content = message,
                        ContentType = BodyType.Html
                    },
                };

                // Trimitem mesajul în canalul specificat
                await _graphClient.Teams[_teamId].Channels[_channelId].Messages.PostAsync(chatMessage);

                Console.WriteLine("Message sent successfully.");
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error sending message: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"General error: {ex.Message}");
            }
        }


        public async Task GetAllTeamsAsync()
        {
            try
            {
                // Obține toate echipele
                var teams = await _graphClient.Teams.GetAsync();

                foreach (var team in teams.Value)
                {
                    Console.WriteLine($"Team ID: {team.Id}, Team Name: {team.DisplayName}");

                    // Obține canalele pentru fiecare echipă
                    var channels = await _graphClient.Teams[team.Id].Channels.GetAsync();

                    foreach (var channel in channels.Value)
                    {
                        Console.WriteLine($"\tChannel ID: {channel.Id}, Channel Name: {channel.DisplayName}");
                    }
                }
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting teams or channels: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"General error: {ex.Message}");
            }
        }
        public async Task<string> GetUserIdByEmailAsync(string email)
        {
            try
            {
                var user = await _graphClient.Users[email].GetAsync();
                return user.Id;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting user by email: {ex.Message}");
                return null;
            }
        }

        public async Task CreateTeamAsync(string displayName, string description)
        {
            try
            {
                var userId = await GetUserIdByEmailAsync("bogdan.alexandru@rlvmgt.onmicrosoft.com");

                if (string.IsNullOrEmpty(userId))
                {
                    Console.WriteLine("User not found.");
                    return;
                }

                var requestBody = new Team
                {
                    DisplayName = displayName,
                    Description = description,
                    Members = new List<ConversationMember>
                    {
                        new AadUserConversationMember
                        {
                            OdataType = "#microsoft.graph.aadUserConversationMember",
                            Roles = new List<string>
                            {
                                "owner",
                            },
                            AdditionalData = new Dictionary<string, object>
                            {
                                { "user@odata.bind", $"https://graph.microsoft.com/v1.0/users('{userId}')" },
                            },
                        },
                    },
                    AdditionalData = new Dictionary<string, object>
                    {
                        { "template@odata.bind", "https://graph.microsoft.com/v1.0/teamsTemplates('standard')" },
                    },
                };

                Team? team = await _graphClient.Teams.PostAsync(requestBody);
                if (team != null)
                {
                    Console.WriteLine($"Team created successfully with ID: {team.Id}");
                }
                else
                {
                    Console.WriteLine("Error creating team: Response is null.");
                }

            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error creating team: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"General error: {ex.Message}");
            }
        }

        public async Task CreateChannelAsync(string channelName, string channelDescription)
        {
            try
            {
                var requestBody = new Channel
                {
                    DisplayName = channelName,
                    Description = channelDescription,
                    MembershipType = ChannelMembershipType.Standard, // sau HiddenMembershipType, în funcție de preferințe
                };

                var channel = await _graphClient.Teams[_teamId].Channels.PostAsync(requestBody);

                Console.WriteLine($"Channel '{channel.DisplayName}' created successfully with ID: {channel.Id}");
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error creating channel: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"General error: {ex.Message}");
            }
        }

    }
}

