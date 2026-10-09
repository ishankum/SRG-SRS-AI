using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using Azure.Storage.Sas;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Net;
using System.Threading.Tasks;

namespace SrsAi.Functions
{
    public class SrsListFunction
    {
        private readonly ILogger _logger;
        private readonly IConfiguration _configuration;
        private readonly BlobServiceClient _blobServiceClient;

        public SrsListFunction(ILoggerFactory loggerFactory, IConfiguration configuration)
        {
            _logger = loggerFactory.CreateLogger<SrsListFunction>();
            _configuration = configuration;

            string connectionString = _configuration["AzureWebJobsStorage"] ?? "";
            _blobServiceClient = new BlobServiceClient(connectionString);
        }

        [Function("GetSrsDrafts")]
        public async Task<HttpResponseData> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "srs/drafts")] HttpRequestData req)
        {
            _logger.LogInformation("Received request to list SRS draft documents with SAS tokens.");

            try
            {
                var draftContainer = _blobServiceClient.GetBlobContainerClient("srs-drafts");
                bool exists = await draftContainer.ExistsAsync();

                var draftItems = new List<object>();

                if (exists)
                {
                    await foreach (BlobItem blobItem in draftContainer.GetBlobsAsync())
                    {
                        var blobClient = draftContainer.GetBlobClient(blobItem.Name);

                        string sasUrl = string.Empty;

                        if (blobClient.CanGenerateSasUri)
                        {
                            BlobSasBuilder sasBuilder = new BlobSasBuilder
                            {
                                BlobContainerName = "srs-drafts",
                                BlobName = blobItem.Name,
                                Resource = "b",
                                ExpiresOn = DateTimeOffset.UtcNow.AddHours(24)
                            };
                            sasBuilder.SetPermissions(BlobSasPermissions.Read);

                            sasUrl = blobClient.GenerateSasUri(sasBuilder).ToString();
                        }
                        else
                        {
                            sasUrl = blobClient.Uri.ToString();
                        }

                        draftItems.Add(new
                        {
                            fileName = blobItem.Name,
                            contentType = blobItem.Properties.ContentType ?? "application/octet-stream",
                            sizeBytes = blobItem.Properties.ContentLength,
                            createdOn = blobItem.Properties.CreatedOn,
                            lastModified = blobItem.Properties.LastModified,
                            downloadSasUrl = sasUrl
                        });
                    }
                }

                var response = req.CreateResponse(HttpStatusCode.OK);
                await response.WriteAsJsonAsync(new
                {
                    success = true,
                    container = "srs-drafts",
                    totalFiles = draftItems.Count,
                    sasTokenValidityHours = 24,
                    files = draftItems
                });

                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error listing SRS drafts: {ex.Message}");

                var errorResponse = req.CreateResponse(HttpStatusCode.InternalServerError);
                await errorResponse.WriteAsJsonAsync(new
                {
                    success = false,
                    error = ex.Message
                });

                return errorResponse;
            }
        }
    }
}
