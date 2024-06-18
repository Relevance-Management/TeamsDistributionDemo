using Microsoft.Graph.Models;
using System.Threading.Tasks;

namespace TeamsDemo.Services.Interfaces
{
    public interface ITeamsService
    {
        Task SendMessageAsync(string message);
        Task GetAllTeamsAsync();
        Task CreateTeamAsync(string displayName, string description);

        Task CreateChannelAsync(string channelName, string channelDescription);
    }
}
