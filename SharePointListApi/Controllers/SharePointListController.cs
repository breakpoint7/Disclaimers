using Microsoft.AspNetCore.Mvc;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json;


// This is a sample Web API controller that returns a list of disclaimers from a SharePoint list.
// A fetch call is used from the office taskpane javascript code to call this controller.  The controller builds a Graph API request to fetch items from a SharePoint list and includes credentials to authenticate the request.
// Requirements:
//  To use this approach, you would need to configure an App Registration in your Azure tenant and grant the necessary permissions to access the SharePoint site.
//  You will need an Application (client) ID, Tenant ID, and Client Secret to authenticate the request.
//  You will need to grant application level permissions to Graph API for Sites.Read.All (or equivalent) to read the SharePoint List.
//  In addition, you need to provide the SharePoint site ID and list ID to fetch the list items.
//  The SharePoint list is expected to use a list columns where "Title" is mapped to "Description", "Text" is mapped to "Text", and "Ver" is mapped to "Version".
// Notes:
// Version is not used, but provided in the event you need to version responses (you would need to add this code)

namespace SharePointListApi.Controllers
{

    public class ListItemFields
    {
        public required string Description { get; set; }

        public required string Text { get; set; }

        public required string Version { get; set; }
    }


    [Route("api/[controller]")]
    [ApiController]
    public class SharePointListController : ControllerBase
    {

        private readonly IHttpClientFactory _clientFactory;

        public SharePointListController(IHttpClientFactory clientFactory)
        {
            _clientFactory = clientFactory;
        }

        [HttpGet]
        public async Task<IActionResult> Get()
        {
            try
            {
                var tenantId = "<tenant id>";  // Tenant where this app is registered
                var clientId = "<app id>";  // Application (client) ID of the registered app
                var clientSecret = "<app secret>";  // Secret key of the registered app (keep this secure and do not hard code this value in your code -- this is for demo purposes only)

                // SharePoint site and list IDs can be either GUIDs or names
                // See https://learn.microsoft.com/en-us/graph/api/list-get?view=graph-rest-1.0&tabs=http and https://learn.microsoft.com/en-us/graph/api/resources/site?view=graph-rest-1.0#id-property
                // Grpah Explore may be helpful in validating these calls if you have problems with the Graph Query (get it working there first, and ensure the Graph call below reflects this complete URL)
                var siteId = "<sharepoint site name or site id>";
                var listId = "<list name or list id>";

                // Get an access token from the registered app to use the SharePoint Graph API
                var token = await GetAccessToken(tenantId, clientId, clientSecret);
                // Retrieve items from the SharePoint list
                var listItems = await GetSharePointListItems(token, siteId, listId);

                return Ok(listItems);
            }
            catch (Exception ex)
            {
                // Log the exception details or handle them as needed
                return StatusCode(500, $"An error occurred while fetching the list items: {ex.Message}");
            }
        }

        private async Task<string> GetAccessToken(string tenantId, string clientId, string clientSecret)
        {
            try
            {
                var client = _clientFactory.CreateClient();

                var request = new HttpRequestMessage(HttpMethod.Post, $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token");
                request.Content = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("client_id", clientId),
                    new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/.default"),
                    new KeyValuePair<string, string>("client_secret", clientSecret),
                    new KeyValuePair<string, string>("grant_type", "client_credentials"),
                });

                var response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var responseStream = await response.Content.ReadAsStringAsync();
                var responseObject = JsonConvert.DeserializeObject<dynamic>(responseStream);

                return responseObject.access_token;
            }
            catch (HttpRequestException httpEx)
            {
                // Log the exception details or handle them as needed
                throw new Exception($"An error occurred while fetching the access token: {httpEx.Message}", httpEx);
            }
            catch (Exception ex)
            {
                // Handle other exceptions
                throw new Exception($"An unexpected error occurred: {ex.Message}", ex);
            }
        }

        private async Task<IActionResult> GetSharePointListItems(string accessToken, string siteId, string listId)
        {
            try
            {
                var client = _clientFactory.CreateClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var requestUrl = $"https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/items?expand=fields(select=Title,Text,Ver)";

                var response = await client.GetAsync(requestUrl);
                response.EnsureSuccessStatusCode();

                var responseStream = await response.Content.ReadAsStringAsync();
                var responseObject = JsonConvert.DeserializeObject<dynamic>(responseStream);

                // Initialize a list to hold the simplified list item fields
                var items = new List<ListItemFields>();

                // Iterate through each item in the response and extract the fields
                // If the columns of your SharePoint list are different, you will need to adjust the fields accordingly
                foreach (var item in responseObject.value)
                {
                    var fields = new ListItemFields
                    {
                        // Map the fields from your response to the properties, assuming `item.fields` contains them
                        Text = item.fields.Text,
                        Description = item.fields.Title,
                        Version = item.fields.Ver
                    };

                    items.Add(fields);
                }

                // Return the JSON string as a ContentResult
                return Ok(items);
            }
            catch(HttpRequestException httpEx)
            {
                // Log the exception details or handle them as needed
                throw new Exception($"An error occurred while fetching the list items: {httpEx.Message}", httpEx);
            }
            catch (Exception ex)
            {
                // Handle other exceptions
                throw new Exception($"An unexpected error occurred: {ex.Message}", ex);
            }   
        }
    }
}
