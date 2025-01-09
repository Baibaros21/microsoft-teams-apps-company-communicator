// <copyright file="StoreMessageActivity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Extensions.DurableTask;
    using Microsoft.Extensions.Caching.Memory;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AdaptiveCard;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Blob;

    /// <summary>
    /// Stores the message in sending notification data table.
    /// </summary>
    public class StoreMessageActivity
    {
        private static readonly string CachePrefixImage = "image_";
        private readonly ISendingNotificationDataRepository sendingNotificationDataRepository;
        private readonly AdaptiveCardCreator adaptiveCardCreator;
        private readonly IMemoryCache memoryCache;
        private readonly IBlobStorageProvider blobStorageProvider;
        private readonly string DEFAULT_LOGO_BLOB_NAME = "DEFAULT_LOGO";
        private readonly string DEFAULT_BANNER_BLOB_NAME = "DEFAULT_BANNER";

        /// <summary>
        /// Initializes a new instance of the <see cref="StoreMessageActivity"/> class.
        /// </summary>
        /// <param name="notificationRepo">Sending notification data repository.</param>
        /// <param name="cardCreator">The adaptive card creator.</param>
        /// <param name="memoryCache">The memory cache.</param>
        /// <param name="blobStorageProvider"></param>
        public StoreMessageActivity(
            ISendingNotificationDataRepository notificationRepo,
            AdaptiveCardCreator cardCreator,
            IMemoryCache memoryCache,
            IBlobStorageProvider blobStorageProvider)
        {
            this.sendingNotificationDataRepository = notificationRepo ?? throw new ArgumentNullException(nameof(notificationRepo));
            this.adaptiveCardCreator = cardCreator ?? throw new ArgumentNullException(nameof(cardCreator));
            this.memoryCache = memoryCache ?? throw new ArgumentNullException(nameof(memoryCache));
            this.blobStorageProvider = blobStorageProvider ?? throw new ArgumentException(nameof(blobStorageProvider));
        }

        /// <summary>
        /// Stores the message in sending notification data table.
        /// </summary>
        /// <param name="notification">A notification to be sent to recipients.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <param name="log">Logging service.</param>
        [FunctionName(FunctionNames.StoreMessageActivity)]
        public async Task RunAsync(
            [ActivityTrigger] NotificationDataEntity notification,
            ILogger log)
        {
            if (notification == null)
            {
                throw new ArgumentNullException(nameof(notification));
            }

            /*            // In case we have blob name instead of URL to public image.
                        if (!string.IsNullOrEmpty(notification.ImageBase64BlobName)
                            && notification.ImageLink.StartsWith(Constants.ImageBase64Format))
                        {
                            var cacheKeySentImage = CachePrefixImage + notification.Id + "-image";
                            bool isCacheEntryExists = this.memoryCache.TryGetValue(cacheKeySentImage, out string imageLink);

                            if (!isCacheEntryExists)
                            {
                                imageLink = await this.sendingNotificationDataRepository.GetImageAsync(notification.ImageBase64BlobName);
                                this.memoryCache.Set(cacheKeySentImage, imageLink, TimeSpan.FromHours(Constants.CacheDurationInHours));

                                log.LogInformation($"Successfully cached the image." +
                                                $"\nNotificationId Id: {notification.Id}");
                            }

                            notification.ImageLink += imageLink;
                        }

                        // In case we have blob name instead of URL to public image.
                        if (!string.IsNullOrEmpty(notification.PosterBase64BlobName)
                            && notification.PosterLink.StartsWith(Constants.ImageBase64Format))
                        {
                            var cacheKeySentImage = CachePrefixImage + notification.Id + "-poster";
                            bool isCacheEntryExists = this.memoryCache.TryGetValue(cacheKeySentImage, out string posterLink);

                            if (!isCacheEntryExists)
                            {
                                posterLink = await this.sendingNotificationDataRepository.GetImageAsync(notification.PosterBase64BlobName);
                                this.memoryCache.Set(cacheKeySentImage, posterLink, TimeSpan.FromHours(Constants.CacheDurationInHours));

                                log.LogInformation($"Successfully cached the image." +
                                                $"\nNotificationId Id: {notification.Id}");
                            }

                            notification.PosterLink += posterLink;
                        }


                        var logoLink = "data:image/jpeg;base64," + await this.blobStorageProvider.DownloadBase64ImageAsync(this.DEFAULT_LOGO_BLOB_NAME);
                        var bannerLink = "data:image/jpeg;base64," + await this.blobStorageProvider.DownloadBase64ImageAsync(this.DEFAULT_BANNER_BLOB_NAME);

                        var defaults = new DefaultsDataEntity
                        {
                            BannerFileName = this.DEFAULT_BANNER_BLOB_NAME,
                            BannerLink = bannerLink,
                            LogoFileName = this.DEFAULT_LOGO_BLOB_NAME,
                            LogoLink = logoLink,
                        };
                        *//* var serializedContent = this.adaptiveCardCreator.CreateAdaptiveCard(notification, defaults).ToJson();

                         // Save Adaptive Card with data uri into blob storage. Blob name = notification.Id.
                         await this.sendingNotificationDataRepository.SaveAdaptiveCardAsync(notification.Id, serializedContent);

                         var serializedContent = await this.sendingNotificationDataRepository.GetAdaptiveCardAsync(notification.Id);*/

            var sendingNotification = new SendingNotificationDataEntity
            {
                PartitionKey = NotificationDataTableNames.SendingNotificationsPartition,
                RowKey = notification.RowKey,
                NotificationId = notification.Id,
                Content = notification.Id,
            };

            await this.sendingNotificationDataRepository.CreateOrUpdateAsync(sendingNotification);
        }


    }


}
