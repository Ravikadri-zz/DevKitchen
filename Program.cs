using Microsoft.Bot.Builder;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

//Ref:https://www.vrdmn.com/search?q=proactive

namespace ConsoleApp2
{
    class Program
    {
        public static async Task Main(string[] args)
        {
            //Teams internal id
            string teamInternalId = "19:bfe3b083bef2488c800c89786b9f01fd@thread.skype";

           
            string serviceUrl = "https://smba.trafficmanager.net/emea/";

            //the upn of the user who should recieve the personal message
            string mentionUserPrincipalName = "reader@dev.onmicrosoft.com";

            //Office 365/Azure AD tenant id for the 
            string tenantId = "";

            //From the Bot Channel Registration
            string botClientID = "bbfc3f45-1a8d-47eb-99f5-dca58ac8a46d";
            string botClientSecret = "";

            var connectorClient = new ConnectorClient(new Uri(serviceUrl), new MicrosoftAppCredentials(botClientID, botClientSecret));

        

            dynamic jsonObject = new JObject();
            //jsonObject.givenName = "Ravi";
            //jsonObject.surname = "Chandra";
            //jsonObject.email = "r@vihandev123.onmicrosoft.com";
            //jsonObject.userPrincipalName = "r@vihandev123.onmicrosoft.comm";
            //jsonObject.tenantId = "a7b57368-7e7b-4r5r5b-b3c5-e3ec26cc53der";
            //jsonObject.userRole = "user";


            jsonObject.givenName = "Reader";
            jsonObject.surname = "1";
            jsonObject.email = "reader@dev1.onmicrosoft.com";
            jsonObject.userPrincipalName = "reader@dev1.onmicrosoft.com";
            jsonObject.tenantId = "a7b5tg56368-7essss7b-4e76-b3c5-e3ec26c45tr3de";
            jsonObject.userRole = "user";


            //Manually fetch user details. Recipient user doesnt have to be in any Team
            //var user = new ChannelAccount {
            //    AadObjectId = "624d2b21-cfc6-40af-b757-3acfed9cb4a7", 
            //    Id = "29:1KL4qDZSwZ7T-5U4b8ghA7cBNdSLanItMvVtIp_HhgJlh_ZLaPTbos7T5hLbDYVwzEcCLpt7dvZk1fNpBo0l8yw", 
            //    Name = "Reader1", 
            //    Properties = jsonObject, 
            //    Role = "null" 
            //};

            //Automatically fetch user details.  User has to memember of the Team where BOT installed. Teams internal id. Uncomment to use it.
            var user = await ((Microsoft.Bot.Connector.Conversations)connectorClient.Conversations).GetConversationMemberAsync(mentionUserPrincipalName, teamInternalId, default);


            var personalMessageActivity = MessageFactory.Text($"Ravis Skynet is active. Machine is taking control!");

            var conversationParameters = new ConversationParameters()
            {
              
                ChannelData = new TeamsChannelData
                {
                    Tenant = new TenantInfo
                    {
                        Id = tenantId,
                    }
                },
                Members = new List<ChannelAccount>() { user },
               
            };



          
            var response = await connectorClient.Conversations.CreateConversationAsync(conversationParameters);

            await connectorClient.Conversations.SendToConversationAsync(response.Id, personalMessageActivity);
        }
    }
}
