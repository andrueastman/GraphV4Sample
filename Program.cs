extern alias BetaLib;
using Beta = BetaLib.Microsoft.Graph;
using Azure.Identity;
using Microsoft.Graph;
using System;
using System.Threading.Tasks;
using System.Net;
using System.Text.Json;

namespace GraphV4Sample
{
    class Program
    {
        public static async Task Main(string[] _)
        {
            GraphServiceClient graphServiceClient = GetAuthenticatedGraphServiceClient();
            // Serialization demo
            await SerilizationDemo(graphServiceClient);
            // List user info
            await GraphResponseDemo(graphServiceClient);

        }

        /// <summary>
        /// Demo of authentication changes
        /// </summary>
        /// <returns></returns>
        public static GraphServiceClient GetAuthenticatedGraphServiceClient()
        {
            // Other TokenCredentials examples are available at https://github.com/microsoftgraph/msgraph-sdk-dotnet/blob/dev/docs/tokencredentials.md
            string[] scopes = new[] { "User.Read", "User.ReadWrite" };
            InteractiveBrowserCredentialOptions interactiveBrowserCredentialOptions = new InteractiveBrowserCredentialOptions()
            {
                ClientId = "CLIENT_ID"
            };
            InteractiveBrowserCredential interactiveBrowserCredential = new InteractiveBrowserCredential(interactiveBrowserCredentialOptions);
            // GraphServiceClient constructor accepts tokenCredential
            GraphServiceClient graphClient = new GraphServiceClient(interactiveBrowserCredential, scopes);
            return graphClient;
        }

        /// <summary>
        /// Demo of changes in serialization
        /// </summary>
        /// <param name="graphServiceClient">The <see cref="GraphServiceClient"/> use</param>
        /// <returns></returns>
        private static async Task SerilizationDemo(GraphServiceClient graphServiceClient)
        {
            // GET user info the normal way
            var user = await graphServiceClient.Me.Request().GetAsync();
            Console.WriteLine("Display Name: " + user.DisplayName);

            // Reading values from the Additional Data
            if (user.AdditionalData.TryGetValue("@odata.context", out object oDataContext))
            {
                string context = ((JsonElement)oDataContext).GetString();
                Console.WriteLine("OData Context: " + context);
            }

            // using the inbuilt serializer to serialize
            var serializedPayload = graphServiceClient.HttpProvider.Serializer.SerializeObject(user);
            // using the inbuilt serializer to deserailize
            var deserializedObject = graphServiceClient.HttpProvider.Serializer.DeserializeObject<User>(serializedPayload);
        }

        /// <summary>
        /// Demo of changes in the GraphResponse
        /// </summary>
        /// <param name="graphServiceClient">The <see cref="GraphServiceClient"/> use</param>
        /// <returns></returns>
        public static async Task GraphResponseDemo(GraphServiceClient graphServiceClient)
        {
            // GET user info the normal way
            Console.WriteLine("Fetching user info ...");
            var user = await graphServiceClient.Me.Request().GetAsync();// No headers/status code in additional data
            Console.WriteLine("Display Name: " + user.DisplayName);
            Console.WriteLine(Environment.NewLine);

            // User object pulled from https://docs.microsoft.com/en-us/graph/api/user-update?view=graph-rest-1.0&tabs=csharp#request
            user.OfficeLocation = "181/21111";

            var userResponse = await graphServiceClient.Me
                .Request()
                .UpdateResponseAsync(user);

            // Check if the status code is as expected
            Console.WriteLine(userResponse.StatusCode == HttpStatusCode.NoContent
                ? "User Updated successfully"
                : "Failed to update user. Status code is not a 204: No Content");

            //var responseHeaders = userResponse.HttpHeaders;
        }
    }
}
