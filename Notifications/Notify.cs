using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Teams;
using Microsoft.Bot.Connector.Teams.Models;
using System;
using System.Web;
using System.Configuration;
using System.Threading.Tasks;
using System.Linq;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace TeamsBot.Notifications
{
    public class Notify
    {
        //Need to replace teamId with Assignment GroupId and do the lookup to get the TeamId
        public static async Task<ResourceResponse> Send1To1MessageToUser(string userAADID, string teamId, string tenantId, 
                                                                         string subEntityId, string messageToSend)
        {
            //we have an AADID for a User in a Team.  
            //Get the users for the Team
            TeamsChannelAccount[] members = await GetMemeberOfTeamAsync(teamId, tenantId);

            string userTeamId = members.Where(x => x.ObjectId == userAADID)
                                       .Select(x => x.Id)                                       
                                       .First();

            if (userTeamId == null) throw new ArgumentException("User not found in team", userAADID);

            ChannelAccount toUser = new ChannelAccount(id: userTeamId);
            return await Send1To1MessageToUser(toUser, teamId, tenantId, subEntityId, messageToSend);
        }

        public static async Task<ResourceResponse> Send1To1MessageToUser(ChannelAccount toUser, string teamId, string toTenantId, 
                                                                         string subEntityId, string messageToSend)
        {
            try
            {
                var response = CONNECTOR.Conversations.CreateOrGetDirectConversation(BOT_USER_ID, toUser, toTenantId);
             
                IMessageActivity sendMe = Activity.CreateMessageActivity();
                sendMe.Text = "What is going on";
                sendMe.Type = ActivityTypes.Message;
                sendMe.Summary = "Kevin Is Cool";
                var ff = sendMe.ChannelData;
                ChannelData data = new ChannelData
                {
                    notification = new Notification
                    {
                        alert = true
                    }
                };
                sendMe.ChannelData = data;
                //sendMe.Attachments.Add(getCard(ASSIGNMNET_ENTITY_ID, teamId));

                return await CONNECTOR.Conversations.SendToConversationAsync((Activity)sendMe, response.Id);
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static async Task<ConversationResourceResponse> SendMessageToGeneralChannelOfTeam(string teamId, 
                                                                                                 string subEntityId,
                                                                                                 string title,
                                                                                                 string subtitle,
                                                                                                 string text)
        {
            var channelData = new Dictionary<string, string>();
            channelData["teamsChannelId"] = teamId;

        
          
            IMessageActivity newMessage = Activity.CreateMessageActivity();
            newMessage.Type = ActivityTypes.Message;
            newMessage.Attachments.Add(getCard(ASSIGNMNET_ENTITY_ID, teamId, title, subtitle, text));

            ConversationParameters conversationParams = new ConversationParameters(isGroup: true,
                                                                                   bot: null,
                                                                                   members: null,
                                                                                   topicName: "AssignmentBot",
                                                                                   activity: (Activity)newMessage,
                                                                                   channelData: channelData);
            return await CONNECTOR.Conversations.CreateConversationAsync(conversationParams);
        }

        public static async Task<ResourceResponse> UpdateMessageToGeneralChannelOfTeam(string chanelId,
                                                                                       string activityId,
                                                                                       string teamId,
                                                                                       string title,
                                                                                       string subtitle,
                                                                                       string text)
        {
          
            IMessageActivity updateMessage = Activity.CreateMessageActivity();
            updateMessage.Type = ActivityTypes.Message;
            updateMessage.Attachments.Add(getCard(ASSIGNMNET_ENTITY_ID, teamId, title, subtitle, text));
            try
            {
                return await CONNECTOR.Conversations.UpdateActivityAsync(chanelId, activityId, (Activity)updateMessage);
            } catch (Exception e)
            {
                throw e;
            }
        }

        public static async Task<TeamsChannelAccount[]> GetMemeberOfTeamAsync(string teamId, string tenantId)
        {
            return await CONNECTOR.Conversations.GetTeamsConversationMembersAsync(teamId, tenantId);
        }


        private static Attachment getCard(string subEntityId, string teamId, string title, string subTitle, string text)
        {
            string deepLink = GetAssignmentDeepLink(subEntityId, teamId);
            var card = new ThumbnailCard
            {
                Title = title,
                Subtitle = subTitle,
                Text = text,
                Images = new List<CardImage>(),
                Buttons = new List<CardAction>
                {
                    new CardAction(type: ActionTypes.OpenUrl, title: "Go To Assignment", 
                                   value: deepLink)
                }
            };
            return card.ToAttachment();
        }

        private static string GetAssignmentDeepLink(string subEntityId, string teamId)
        {
            //first we need to create a Context for the Link
            //the channelId for our Assignments tab is the TeamID as that is the
            //channelId for the General Tab
            Context context = new Context
            {
                subEntityId = subEntityId,
                channelId = teamId
            };

            return ASSIGNMENT_DEEP_LINK_URL + HttpUtility.UrlEncode(ASSIGNMNET_APP_ID) +
                   "/" + HttpUtility.UrlEncode(ASSIGNMNET_ENTITY_ID) + "?context="
                   + HttpUtility.UrlEncode(JsonConvert.SerializeObject(context));


        }

        private static ChannelAccount BOT_USER_ID = new ChannelAccount(id: ConfigurationManager.AppSettings["BotGUID"].ToString());
        private static ConnectorClient CONNECTOR = new ConnectorClient(new Uri(ConfigurationManager.AppSettings["BotFrameWorkURI"].ToString()),
                                                                       ConfigurationManager.AppSettings["MicrosoftAppId"].ToString(),
                                                                       ConfigurationManager.AppSettings["MicrosoftAppPassword"].ToString());

        private static string ASSIGNMNET_APP_ID = "88aa3ede-1d0a-4dd9-af10-013c194420aa";
        private static string ASSIGNMNET_ENTITY_ID = "entityId";
        private static string ASSIGNMENT_DEEP_LINK_URL = "https://teams.microsoft.com/l/entity/";

    }

    class Context
    {
        public string subEntityId { get; set; }
        public string channelId { get; set; }
    }

    class ChannelData
    {
        public Notification notification { get; set; }
    }

    class Notification
    {
        public Boolean alert { get; set; }
    }
}