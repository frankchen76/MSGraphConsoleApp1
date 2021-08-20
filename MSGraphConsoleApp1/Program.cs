using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace MSGraphConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] scopes = { "User.Read" };

            InteractiveBrowserCredentialOptions interactiveBrowserCredentialOptions = new InteractiveBrowserCredentialOptions()
            {
                ClientId = clientId
            };
            InteractiveBrowserCredential interactiveBrowserCredential = new InteractiveBrowserCredential(interactiveBrowserCredentialOptions);

            GraphServiceClient graphClient = new GraphServiceClient(myBrowserCredential, scopes); // you can pass the TokenCredential directly to the GraphServiceClient

            User me = await graphClient.Me.Request()
                            .GetAsync();
        }
    }
}
