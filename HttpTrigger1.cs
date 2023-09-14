using System.Net;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using PnP.Core.Services;
using Newtonsoft.Json;
using PnP.Core.QueryModel;
using AngleSharp.Dom;
using PnP.Core.Model.SharePoint;
using Microsoft.Extensions.Primitives;
using Azure.Storage.Blobs;
using System.IO;
using System.Threading.Tasks;
using Azure.Storage.Blobs.Models;
using PnP.Core.Model.Teams;
using AngleSharp.Common;


// -=-
// SharePoint PnP Library Docs
// https://pnp.github.io/pnpcore/using-the-sdk/readme.html

namespace ClipWizzard
{

    public static class ChangeTokenStore
    {
        public static PnP.Core.Model.SharePoint.IChangeToken Token { get; set; }
    }

    // supporting classes
    public class ResponseModel<T>
    {
        [JsonProperty(PropertyName = "value")]
        public List<T> Value { get; set; }
    }
    public class NotificationModel
    {
        [JsonProperty(PropertyName = "subscriptionId")]
        public string SubscriptionId { get; set; }

        [JsonProperty(PropertyName = "clientState")]
        public string ClientState { get; set; }

        [JsonProperty(PropertyName = "expirationDateTime")]
        public DateTime ExpirationDateTime { get; set; }

        [JsonProperty(PropertyName = "resource")]
        public string Resource { get; set; }

        [JsonProperty(PropertyName = "tenantId")]
        public string TenantId { get; set; }

        [JsonProperty(PropertyName = "siteUrl")]
        public string SiteUrl { get; set; }

        [JsonProperty(PropertyName = "webId")]
        public string WebId { get; set; }
    }
    public class SubscriptionModel
    {
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string Id { get; set; }

        [JsonProperty(PropertyName = "clientState", NullValueHandling = NullValueHandling.Ignore)]
        public string ClientState { get; set; }

        [JsonProperty(PropertyName = "expirationDateTime")]
        public DateTime ExpirationDateTime { get; set; }

        [JsonProperty(PropertyName = "notificationUrl")]
        public string NotificationUrl { get; set; }

        [JsonProperty(PropertyName = "resource", NullValueHandling = NullValueHandling.Ignore)]
        public string Resource { get; set; }
    }
    public class HttpTrigger1
    {
        private readonly ILogger _logger;
        private readonly IPnPContextFactory contextFactory;
        public HttpTrigger1(ILoggerFactory loggerFactory, IPnPContextFactory pnpContextFactory)
        {
            _logger = loggerFactory.CreateLogger<HttpTrigger1>();
            contextFactory = pnpContextFactory;
        }

        [Function("HttpTrigger1")]
        public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
        {
            
            string connectionString =  Environment.GetEnvironmentVariable("StorageConnectionString");
            string containerName =  Environment.GetEnvironmentVariable("StorageContainerName");

            _logger.LogInformation("C# HTTP trigger function processed a request.");
            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "text/plain; charset=utf-8");
            response.WriteString("Welcome to Azure Functions!");

            using (var context = await contextFactory.CreateAsync("Default"))
            {
                try
                {
                    var web = await context.Web.GetAsync();
                    var list = await context.Web.Lists.GetByTitleAsync("Documents", p => p.Title, p => p.Items);
                    foreach (var listItem in list.Items.AsRequested())
                    {
                        // Load list item
                        var item = await list.Items.GetByIdAsync(listItem.Id, li => li.All, li => li.File, li => li.FieldValuesAsText, li => li.Title);
                        byte[] downloadedContentBytes = await item.File.GetContentBytesAsync();

                        // Iterate through the item and selectivly pick the metadata we want to add to the blob
                        // sadly, dumping all the feilds hits the 8kb limit of Blob metadata and causes the API 
                        // call to fail
                        var metadata = new Dictionary<string, string>();
                        foreach (var kvp in item.FieldValuesAsText.Values)
                        {
                            if (kvp.Key.ToString().ToLower().Contains("id") ||
                                kvp.Key.ToString().ToLower().Contains("userrole") ||
                                kvp.Key.ToString().ToLower().Contains("url"))
                            {
                                metadata[kvp.Key.ToString()] = kvp.Value.ToString();
                            }
                        }

                        // Set up a blob client, and add the metadata
                        var blobServiceClient = new BlobServiceClient(connectionString);
                        var containerClient = blobServiceClient.GetBlobContainerClient(containerName);
                        var blobClient = containerClient.GetBlobClient(item.File.Name);
                        var options = new BlobUploadOptions
                        {
                            Metadata = metadata
                        };

                        // Upload the file to the blob
                        await blobClient.UploadAsync(new MemoryStream(downloadedContentBytes), options, new CancellationToken());

                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex.Message);
                }
            }
            return response;
        }

        // [Function("HttpTrigger1")]
        // public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
        // {
        //     _logger.LogInformation("C# HTTP trigger function processed a request.");

        //     var response = req.CreateResponse(HttpStatusCode.OK);
        //     response.Headers.Add("Content-Type", "text/plain; charset=utf-8");
        //     response.WriteString("Welcome to Azure Functions!");


        //     // Grab the validationToken URL parameter
        //     string validationToken = req.Query["validationtoken"];

        //     // If a validation token is present, we need to respond within 5 seconds by
        //     // returning the given validation token. This only happens when a new
        //     // webhook is being added
        //     if (validationToken != null)
        //     {
        //         var r = req.CreateResponse(HttpStatusCode.OK);
        //         r.Headers.Add("Content-Type", "text/plain; charset=utf-8");
        //         r.WriteString(validationToken);
        //         return r;
        //     }

        //     _logger.LogInformation($"SharePoint triggered our webhook...great :-)");
        //     var content = await new StreamReader(req.Body).ReadToEndAsync();
        //     _logger.LogInformation($"Received following payload: {content}");
        //     if (content.Length > 0)
        //     {

        //         var notifications = JsonConvert.DeserializeObject<ResponseModel<NotificationModel>>(content).Value;
        //         _logger.LogInformation($"Found {notifications.Count} notifications");

        //         if (notifications.Count > 0)
        //         {
        //             _logger.LogInformation($"Processing notifications...");
        //             foreach (var notification in notifications)
        //             {
        //                 // add message to the queue
        //                 string message = JsonConvert.SerializeObject(notification);
        //                 _logger.LogInformation($"Before adding a message to the queue. Message content: {message}");
        //                 using (var context = await contextFactory.CreateAsync("Default"))
        //                 {
        //                     try
        //                     {
        //                         var web = await context.Web.GetAsync();
        //                         // Get all type of changes for a SharePoint list limited to the 50 most recent changes
        //                         var list = await context.Web.Lists.GetByTitleAsync("Documents");

        //                         _logger.LogInformation($"Change Token: {0}", ChangeTokenStore.Token.ToString());

        //                         var changes = await list.GetChangesAsync(new ChangeQueryOptions(false, true)
        //                         {
        //                             List = true,
        //                             ChangeTokenEnd = ChangeTokenStore.Token
        //                         }
        //                         );
        //                         _logger.LogInformation(changes.Count().ToString());
        //                         ChangeTokenStore.Token = changes.Last().ChangeToken;
        //                     }
        //                     catch (Exception ex)
        //                     {
        //                         _logger.LogInformation(ex.Message);
        //                     }
        //                 }
        //                 _logger.LogInformation($"Message added :-)");
        //             }
        //         }
        //         return response;
        //     }
        //     else
        //     {
        //         return response;
        //     }
        // }
    }
}
