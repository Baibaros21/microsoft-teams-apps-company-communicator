// <copyright file="BlobStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Blob
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text.RegularExpressions;
    using System.Threading.Tasks;
    using global::Azure;
    using global::Azure.Storage.Blobs;
    using global::Azure.Storage.Blobs.Models;
    using global::Azure.Storage.Blobs.Specialized;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Clients;

    /// <summary>
    /// Provider for handling Azure Blob Storage operations like uploading/downloading images/adaptive cards from blob.
    /// </summary>
    public class BlobStorageProvider : IBlobStorageProvider
    {
        /// <summary>
        /// prefix for data uri image (png).
        /// </summary>
        public const string ImageBase64FormatPng = "data:image/png;base64,";

        /// <summary>
        /// prefix for data uri image (jpeg).
        /// </summary>
        public const string ImageBase64FormatJpeg = "data:image/jpeg;base64,";

        /// <summary>
        /// prefix for data uri image (gif).
        /// </summary>
        public const string ImageBase64FormatGif = "data:image/gif;base64,";

        /// <summary>
        /// blob container name for serilized sent adaptive cards.
        /// </summary>
        public const string SentCardsBlobContainerName = "sentcards";

        /// <summary>
        /// blob container name for images in base64 format.
        /// </summary>
        public const string ImagesBlobContainerName = "images";

        /// <summary>
        /// blob container name for images in base64 format.
        /// </summary>
        public const string TemplatesBlobContainerName = "templates";

        private readonly IStorageClientFactory storageClientFactory;

        /// <summary>
        /// Instance to send logs to the telemetry service.
        /// </summary>
        private readonly ILogger<BlobStorageProvider> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="BlobStorageProvider"/> class.
        /// </summary>
        /// <param name="storageClientFactory">The storage client factory.</param>
        /// <param name="logger">The logging service.</param>
        public BlobStorageProvider(IStorageClientFactory storageClientFactory, ILogger<BlobStorageProvider> logger)
        {
            this.storageClientFactory = storageClientFactory ?? throw new ArgumentNullException(nameof(storageClientFactory));
            this.logger = logger ?? throw new ArgumentException(nameof(logger));
        }

        /// <inheritdoc/>
        public async Task<string> UploadBase64ImageAsync(string blobName, string base64Image)
        {
            string prefix = null;
            if (base64Image.StartsWith(ImageBase64FormatJpeg))
            {
                prefix = ImageBase64FormatJpeg;
            }
            else if (base64Image.StartsWith(ImageBase64FormatPng))
            {
                prefix = ImageBase64FormatPng;
            }
            else if (base64Image.StartsWith(ImageBase64FormatGif))
            {
                prefix = ImageBase64FormatGif;
            }
            else
            {
                throw new FormatException("Image has unsupported format. Only jpeg, png and gif formats are supported.");
            }

            string[] strings = base64Image.Split(prefix);

            try
            {
                var blobContainerClient = await this.GetBlobContainer(ImagesBlobContainerName);

                var blob = blobContainerClient.GetBlobClient(blobName);
                byte[] imageBytes = Convert.FromBase64String(strings[1]);

                using (Stream stream = new MemoryStream(imageBytes))
                {
                    await blob.UploadAsync(stream, true);
                }

                return prefix;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while uploading image to Azure Blob Storage. Blob name : {blobName}, Error details: {ex.Message}");
                throw;
            }
        }

        /// <inheritdoc/>
        public async Task<string> DownloadBase64ImageAsync(string blobName)
        {
            try
            {
                var blobContainerClient = await this.GetBlobContainer(ImagesBlobContainerName);

                var blob = blobContainerClient.GetBlobClient(blobName);

                byte[] imageBytes = null;
                using (var stream = new MemoryStream())
                {
                    await blob.DownloadToAsync(stream);
                    stream.Position = 0;
                    imageBytes = stream.ToArray();
                }

                return Convert.ToBase64String(imageBytes);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while downloading image from Azure Blob Storage. Image blob container name : {ImagesBlobContainerName}, Error details: {ex.Message}");
                throw;
            }
        }

        /// <inheritdoc/>
        public async Task UploadAdaptiveCardAsync(string blobName, string adaptiveCard)
        {
            try
            {
                var blobContainerClient = await this.GetBlobContainer(SentCardsBlobContainerName);

                var blob = blobContainerClient.GetBlobClient(blobName);
                await blob.DeleteIfExistsAsync(DeleteSnapshotsOption.IncludeSnapshots);

                using (var stream = new MemoryStream())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write(adaptiveCard);
                    writer.Flush();
                    stream.Position = 0;
                    var blobHttpHeader = new BlobHttpHeaders() { ContentType = "application/json" };
                    await blob.UploadAsync(stream, blobHttpHeader);
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while uploading AC to Azure Blob Storage. Blob name : {blobName}, Error details: {ex.Message}");
                throw;
            }
        }


        /// <inheritdoc/>
        public async Task UploadCardTemplateAsync(string blobName, string adaptiveCard)
        {
            try
            {
                var blobContainerClient = await this.GetBlobContainer(TemplatesBlobContainerName);

                var blob = blobContainerClient.GetBlobClient(blobName);
                await blob.DeleteIfExistsAsync(DeleteSnapshotsOption.IncludeSnapshots);

                using (var stream = new MemoryStream())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write(adaptiveCard);
                    writer.Flush();
                    stream.Position = 0;
                    var blobHttpHeader = new BlobHttpHeaders() { ContentType = "application/json" };
                    await blob.UploadAsync(stream, blobHttpHeader);
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while uploading AC to Azure Blob Storage. Blob name : {blobName}, Error details: {ex.Message}");
                throw;
            }
        }

        /// <inheritdoc/>
        public async Task<string> DownloadAdaptiveCardAsync(string blobName)
        {
            try
            {
                var blobContainerClient = await this.GetBlobContainer(SentCardsBlobContainerName);
                var blob = blobContainerClient.GetBlobClient(blobName);

                using (var stream = new MemoryStream())
                using (var reader = new StreamReader(stream))
                {
                    await blob.DownloadToAsync(stream);
                    stream.Position = 0;
                    return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while downloading AC from Azure Blob Storage. Blob name : {blobName} , Error details: {ex.Message}");
                throw;
            }
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<CardTemplatedataEntity>> DownloadAllCardTemplatesAsync()
        {
            try
            {
                var blobContainerClient = await this.GetBlobContainer(TemplatesBlobContainerName);

                var blobs = blobContainerClient.GetBlobs();
                var templates = new List<CardTemplatedataEntity>();
                foreach (var blob in blobs)
                {
                    var client = blobContainerClient.GetBlobClient(blob.Name);
                    using (var stream = new MemoryStream())
                    using (var reader = new StreamReader(stream))
                    {
                        await client.DownloadToAsync(stream);
                        stream.Position = 0;

                        templates.Add(new CardTemplatedataEntity { Name = blob.Name, Card = reader.ReadToEnd() });
                    }
                }

                return templates;

            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while downloading templates from Azure Blob Storage. Blob container name : {TemplatesBlobContainerName} , Error details: {ex.Message}");
                throw;
            }
        }


        /// <inheritdoc/>
        public async Task<string> DownloadCardTemplateAsync(string blobName)
        {
            try
            {
                var blobContainerClient = await this.GetBlobContainer(TemplatesBlobContainerName);

                var blob = blobContainerClient.GetBlobClient(blobName);

                using (var stream = new MemoryStream())
                using (var reader = new StreamReader(stream))
                {
                    await blob.DownloadToAsync(stream);
                    stream.Position = 0;
                    return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while downloading AC from Azure Blob Storage. Blob name : {blobName} , Error details: {ex.Message}");
                throw;
            }
        }

        /// <inheritdoc/>
        public async Task DeleteImageBlobAsync(string blobName)
        {
            await this.DeleteBlobAsync(blobName, ImagesBlobContainerName);
        }

        /// <inheritdoc/>
        public async Task CopyImageBlobAsync(string blobName, string newBlobName)
        {
            await this.CopyBlobAsync(blobName, newBlobName, ImagesBlobContainerName);
        }

        private async Task DeleteBlobAsync(string blobName, string blobContainerName)
        {
            try
            {
                var blobContainerClient = await this.GetBlobContainer(blobContainerName);
                var blob = blobContainerClient.GetBlobClient(blobName);
                await blob.DeleteIfExistsAsync(DeleteSnapshotsOption.IncludeSnapshots);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while deleting blob from Azure Blob Storage. Error details: {ex.Message}");
                throw;
            }
        }

        private async Task CopyBlobAsync(string blobName, string newBlobName, string blobContainerName)
        {
            var blobContainerClient = await this.GetBlobContainer(blobContainerName);
            var sourceBlob = blobContainerClient.GetBlobClient(blobName);

            try
            {
                if (await sourceBlob.ExistsAsync())
                {
                    // Lease the source blob for the copy operation
                    // to prevent another client from modifying it.
                    BlobLeaseClient lease = sourceBlob.GetBlobLeaseClient();

                    // Specifying -1 for the lease interval creates an infinite lease.
                    await lease.AcquireAsync(TimeSpan.FromSeconds(-1));

                    // Get a BlobClient representing the destination blob with a unique name.
                    var destBlob = blobContainerClient.GetBlobClient(newBlobName);

                    // Start the copy operation.
                    var status = await destBlob.StartCopyFromUriAsync(sourceBlob.Uri);
                    await status.WaitForCompletionAsync();

                    // Update the source blob's properties.
                    BlobProperties sourceProperties = await sourceBlob.GetPropertiesAsync();

                    // Break the lease on the source blob.
                    // Update the source blob's properties to check the lease state.
                    if (sourceProperties.LeaseState == LeaseState.Leased)
                    {
                        await lease.BreakAsync();
                        sourceProperties = await sourceBlob.GetPropertiesAsync();
                    }
                }
            }
            catch (RequestFailedException ex)
            {
                this.logger.LogError(ex, $"Error while copying blob in Azure Blob Storage Container.");
                throw;
            }
        }

        private async Task<BlobContainerClient> GetBlobContainer(string blobContainerName)
        {
            try
            {
                var container = this.storageClientFactory.CreateBlobContainerClient(blobContainerName);
                await container.CreateIfNotExistsAsync();
                await container.SetAccessPolicyAsync(PublicAccessType.None);

                return container;
            }
            catch (RequestFailedException ex)
            {
                this.logger.LogError($"Cannot find blob container: {blobContainerName} - error details: {ex.Message}");
                throw;
            }
        }
    }
}