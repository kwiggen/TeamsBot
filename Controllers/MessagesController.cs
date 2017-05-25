using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Teams;
using System;
using System.Configuration;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using TeamsBot.Dialogs;

namespace TeamsBot
{
    [BotAuthentication]
    [TenantFilter]
    public class MessagesController : ApiController
    {

        public MessagesController()
        {
            this.connectorClient = new ConnectorClient(new Uri(ConfigurationManager.AppSettings["BotFrameWorkURI"].ToString()),
                                                       ConfigurationManager.AppSettings["MicrosoftAppId"].ToString(),
                                                       ConfigurationManager.AppSettings["MicrosoftAppPassword"].ToString());
        }


        /// <summary>
        /// POST: api/Messages
        /// Receive a message from a user and reply to it
        /// </summary>
        public async Task<HttpResponseMessage> Post([FromBody]Activity activity)
        {
            return await MessageProcessor.HandleIncomingRequest(activity, this.connectorClient);
        }

        private ConnectorClient connectorClient;

    }
}