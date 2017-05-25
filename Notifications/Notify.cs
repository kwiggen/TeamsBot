using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Teams;
using Microsoft.Bot.Connector.Teams.Models;
using System;
using System.Configuration;
using System.Threading.Tasks;
using System.Linq;
using System.Collections.Generic;

namespace TeamsBot.Notifications
{
    public class Notify
    {
        //Need to replace teamId with Assignment GroupId and do the lookup to get the TeamId
        public static async Task<ResourceResponse> Send1To1MessageToUser(string userAADID, string teamId, string tenantId, string messageToSend)
        {
            //we have an AADID for a User in a Team.  
            //Get the users for the Team
            TeamsChannelAccount[] members = await GetMemeberOfTeamAsync(teamId, tenantId);

            string userTeamId = members.Where(x => x.ObjectId == userAADID)
                                       .Select(x => x.Id)                                       
                                       .First();

            if (userTeamId == null) throw new ArgumentException("User not found in team", userAADID);

            ChannelAccount toUser = new ChannelAccount(id: userTeamId);
            return await Send1To1MessageToUser(toUser, tenantId, messageToSend);
        }

        public static async Task<ResourceResponse> Send1To1MessageToUser(ChannelAccount toUser, string toTenantId, string messageToSend)
        {
            try
            {
                var response = CONNECTOR.Conversations.CreateOrGetDirectConversation(BOT_USER_ID, toUser, toTenantId);
             
                IMessageActivity sendMe = Activity.CreateMessageActivity();
                sendMe.Text = messageToSend;
                sendMe.Type = ActivityTypes.Message;

                return await CONNECTOR.Conversations.SendToConversationAsync((Activity)sendMe, response.Id);
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static async Task<ConversationResourceResponse> SendMessageToGeneralChannelOfTeam(string teamId, string messageToSend)
        {
            var channelData = new Dictionary<string, string>();
            channelData["teamsChannelId"] = teamId;

        
          
            IMessageActivity newMessage = Activity.CreateMessageActivity();
            newMessage.Type = ActivityTypes.Message;
            newMessage.Attachments.Add(getCard());
            //newMessage.Text = messageToSend;

            ConversationParameters conversationParams = new ConversationParameters(isGroup: true,
                                                                                   bot: null,
                                                                                   members: null,
                                                                                   topicName: "AssignmentBot",
                                                                                   activity: (Activity)newMessage,
                                                                                   channelData: channelData);
            return await CONNECTOR.Conversations.CreateConversationAsync(conversationParams);
        }

        public static async Task<TeamsChannelAccount[]> GetMemeberOfTeamAsync(string teamId, string tenantId)
        {
            return await CONNECTOR.Conversations.GetTeamsConversationMembersAsync(teamId, tenantId);
        }


        private static Attachment getCard()
        {
            var card = new ThumbnailCard
            {
                Title = "Here be a Title",
                Subtitle = "Here be a Subtitle",
                Text = "And some text",
                Images = new List<CardImage>(),
                Buttons = new List<CardAction>
                {
                    new CardAction(type: ActionTypes.OpenUrl, title: "Open Microsoft", value: "https://www.microsoft.com/en-us/")
                }
            };
            return card.ToAttachment();
        }
        
        private static ChannelAccount BOT_USER_ID = new ChannelAccount(id: ConfigurationManager.AppSettings["BotGUID"].ToString());
        private static ConnectorClient CONNECTOR = new ConnectorClient(new Uri(ConfigurationManager.AppSettings["BotFrameWorkURI"].ToString()),
                                                                       ConfigurationManager.AppSettings["MicrosoftAppId"].ToString(),
                                                                       ConfigurationManager.AppSettings["MicrosoftAppPassword"].ToString());
    }
}