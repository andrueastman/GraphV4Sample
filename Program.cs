extern alias BetaLib;
using Beta = BetaLib.Microsoft.Graph;
using Azure.Identity;
using Microsoft.Graph;
using System;
using System.Threading.Tasks;
using System.Net;

namespace GraphV4Sample
{
    class Program
    {
        public static async Task Main(string[] _)
        {
            string clientId = "Insert_Client_ID_here";
            string[] scopes = new[] { "User.Read", "User.ReadWrite" };

            // Create the client with the TokenCredential
            InteractiveBrowserCredential interactiveBrowserCredential = new InteractiveBrowserCredential(clientId);
            GraphServiceClient graphServiceClient = new GraphServiceClient(interactiveBrowserCredential, scopes);

            // List user info
            await ListUserInfo(graphServiceClient);

            // Update user
            await UpdateUserWithGraphResponse(graphServiceClient);

            // List user info again to confirm update
            await ListUserInfo(graphServiceClient);

        }

        public static async Task ListUserInfo(GraphServiceClient graphServiceClient)
        {
            // GET user info the normal way
            Console.WriteLine("Fetching user info ...");
            var user = await graphServiceClient.Me.Request().GetAsync();
            Console.WriteLine("Display Name: " + user.DisplayName);
            Console.WriteLine("User Principal Name: " + user.UserPrincipalName);
            Console.WriteLine("Office Location: " + user.OfficeLocation);
            Console.WriteLine(Environment.NewLine);

            // GET user info with graph response.
            Console.WriteLine("Fetching user info (with GraphResponse)...");
            var userResponse = await graphServiceClient.Me.Request().GetResponseAsync();

            // Deserialize the response
            if (userResponse.StatusCode == HttpStatusCode.OK)
            {
                Console.WriteLine(userResponse.StatusCode);
                var userObject = await userResponse.GetResponseObjectAsync();
                Console.WriteLine("Display Name: " + userObject.DisplayName);
                Console.WriteLine("User Principal Name: " + userObject.UserPrincipalName);
                Console.WriteLine("Office Location: " + userObject.OfficeLocation);
                Console.WriteLine(Environment.NewLine);
            }
        }

        public static async Task UpdateUserWithGraphResponse(GraphServiceClient graphServiceClient)
        {
            // User object pulled from https://docs.microsoft.com/en-us/graph/api/user-update?view=graph-rest-1.0&tabs=csharp#request
            var user = new User
            {
                OfficeLocation = "181/21111"
            };

            var userResponse = await graphServiceClient.Me
                .Request()
                .UpdateResponseAsync(user);

            // Check if the status code is as expected
            Console.WriteLine(userResponse.StatusCode == HttpStatusCode.NoContent
                ? "User Updated successfully"
                : "Failed to update user. Status code is not a 204: No Content");

            Console.WriteLine(Environment.NewLine);
        }
    }
}
