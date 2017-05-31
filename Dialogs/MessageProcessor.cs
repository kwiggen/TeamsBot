using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Teams;
using Microsoft.Bot.Connector.Teams.Models;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using TeamsBot.Notifications;

namespace TeamsBot.Dialogs
{
    public class MessageProcessor
    {
        public static async Task<HttpResponseMessage> HandleIncomingRequest(Activity activity, ConnectorClient connector)
        {
            switch (activity.GetActivityType())

            {
                case ActivityTypes.Message:
                    await HandleTextMessage(activity, connector);
                    break;

                case ActivityTypes.ConversationUpdate:
                    HandleConversationUpdates(activity, connector);
                    break;

                case ActivityTypes.Invoke:
                case ActivityTypes.ContactRelationUpdate:
                case ActivityTypes.Typing:
                case ActivityTypes.DeleteUserData:
                case ActivityTypes.Ping:
                default:
                    break;
            }

            return new HttpResponseMessage(HttpStatusCode.OK);

        }

        private static async Task HandleTextMessage(Activity activity, ConnectorClient connector)
        {
            //In V1 this method will do nothing, as we won't respond
            //Below is code to test the Notify lib
           
            string teamId = activity.GetChannelData<TeamsChannelData>().Team.Id;
            string tenantId = activity.GetTenantId();
            var response = await Notify.Send1To1MessageToUser(KEVIN_AAD_ID, teamId, tenantId, "subEntity123", "Hello World 1-1 from AssignmentsBot!!!");
            var response2 = await Notify.SendMessageToGeneralChannelOfTeam(teamId, "subEntity456", "New Conversation from TeamBot");


            //now respond to what the user asked the bot directly
            Activity replyActivity = activity.CreateReply();

            TeamsChannelAccount[] members = await Notify.GetMemeberOfTeamAsync(teamId, tenantId);
            replyActivity.Text = "These are the member userids returned by the GetConversationMembers() function:";

            members.ToList().ForEach(x => replyActivity.Text += ", " + x.GivenName + " = " + x.ObjectId);
            await connector.Conversations.ReplyToActivityAsync(replyActivity);
        }

        private static void HandleConversationUpdates(Activity activity, ConnectorClient connectorClient)
        {
            TeamEventBase eventData = activity.GetConversationUpdateData();

            switch (eventData.EventType)
            {
                case TeamEventType.ChannelCreated:
                    {
                        break;
                    }
                case TeamEventType.ChannelDeleted:
                    {
                        break;
                    }
                case TeamEventType.MembersAdded:
                    {
                        //When a Bot is added to a Team, the MembersAdded is triggered for Adding this Bot to the Team
                        //We need to make sure that our Class has the Team ID needed to send messages
                        MembersAddedEvent memberAddedEvent = eventData as MembersAddedEvent;
                        string teamId = memberAddedEvent.Team.Id;
                        string serviceUrl = activity.ServiceUrl;
                        if (teamId != null) //its possible this is a 1-1 conversation
                        {
                            //TODO: Here is where we would write the teamId and serviceUrl into the Class storage so we have it later
                        }
                        break;
                    }
                case TeamEventType.MembersRemoved:
                    {
                        break;
                    }
            }
        }

        private static string KEVIN_AAD_ID = ConfigurationManager.AppSettings["KWID"].ToString();

    }
}


