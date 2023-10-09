// <copyright file="AdaptiveCardCreator.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AdaptiveCard
{
    using System;
    using System.Collections.Generic;
    using System.Text.Encodings.Web;
    using System.Web;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;

    /// <summary>
    /// Adaptive Card Creator service.
    /// </summary>
    public class AdaptiveCardCreator
    {
        /// <summary>
        /// Creates an adaptive card.
        /// </summary>
        /// <param name="notificationDataEntity">Notification data entity.</param>
        /// <returns>An adaptive card.</returns>
        public virtual AdaptiveCard CreateAdaptiveCard(NotificationDataEntity notificationDataEntity)
        {
            return this.CreateAdaptiveCard(
               title: notificationDataEntity.Title,
               imageUrl: notificationDataEntity.ImageLink,
               summary: notificationDataEntity.Summary,
               author: notificationDataEntity.Author,
               buttonTitle: notificationDataEntity.ButtonTitle,
               buttonUrl: notificationDataEntity.ButtonLink,
               notificationId: notificationDataEntity.Id,
               department: notificationDataEntity.Department,
               posterLink: notificationDataEntity.PosterLink,
               videoLink: notificationDataEntity.VideoLink,
               template: notificationDataEntity.Template

                );
             

        }

        private string UrlEncoder(string videoName, string videoURL, string websiteURL)
        {
            Uri videoUri = new Uri(videoURL);
            bool isSPO = videoUri.Host.Contains(".sharepoint.com");

            if (isSPO)
            {
                string TeamsLogon = "/_layouts/15/teamslogon.aspx?spfx=true&dest=";
                string videoURLSPO = $"https://{videoUri.Host}{TeamsLogon}{videoUri.PathAndQuery}";
                return HttpUtility.UrlEncode($"{{\"contentUrl\":\"{videoURLSPO}\",\"websiteUrl\":\"{websiteURL}\",\"name\":\"{videoName}\"}}");
            }
            else
            {
                return HttpUtility.UrlEncode($"{{\"contentUrl\":\"{videoURL}\",\"websiteUrl\":\"{websiteURL}\",\"name\":\"{videoName}\"}}");
            }
        }


        /// <summary>
        /// Create an adaptive card instance.
        /// </summary>
        /// <param name="title">The adaptive card's title value.</param>
        /// <param name="imageUrl">The adaptive card's image URL.</param>
        /// <param name="summary">The adaptive card's summary value.</param>
        /// <param name="author">The adaptive card's author value.</param>
        /// <param name="buttonTitle">The adaptive card's button title value.</param>
        /// <param name="buttonUrl">The adaptive card's button url value.</param>
        /// <param name="notificationId">The notification id.</param>
        /// <param name="template"></param>
        /// <param name="department"></param>
        /// <param name="posterLink"></param>
        /// <param name="videoLink"></param>
        /// <returns>The created adaptive card instance.</returns>
        public AdaptiveCard CreateAdaptiveCard(
            string title,
            string imageUrl,
            string summary,
            string author,
            string buttonTitle,
            string buttonUrl,
            string notificationId,
            string department ,
            string posterLink, 
            string videoLink,
            string template
            )
        {

                

            string TeamsInternalAppID = "148a66bb-e83d-425a-927d-09f4299a9274"; 

            var version = new AdaptiveSchemaVersion(1, 0);
            AdaptiveCard card = new AdaptiveCard(version);

            if (!string.IsNullOrEmpty(department))
            {
                card.Body.Add(new AdaptiveTextBlock()
                {
                    Text = department,
                    Size = AdaptiveTextSize.Medium,
                    Weight = AdaptiveTextWeight.Bolder,
                    Wrap = true,
                    HorizontalAlignment =AdaptiveHorizontalAlignment.Center,
                });
            }

            card.Body.Add(new AdaptiveTextBlock()
            {
                Text = title,
                Size = AdaptiveTextSize.ExtraLarge,
                Weight = AdaptiveTextWeight.Bolder,
                Wrap = true,
            });


            if (!string.IsNullOrWhiteSpace(imageUrl))
            {

                Console.WriteLine(imageUrl);
                var img = new AdaptiveImageWithLongUrl()
                {
                    LongUrl = imageUrl,
                    Spacing = AdaptiveSpacing.Default,
                    Size = AdaptiveImageSize.Stretch,
                    AltText = string.Empty,
                };

                // Image enlarge support for Teams web/desktop client.
                img.AdditionalProperties.Add("msteams", new { AllowExpand = true });

                card.Body.Add(img);
            }

            if (!string.IsNullOrWhiteSpace(videoLink) && !string.IsNullOrWhiteSpace(posterLink))
            {
                string videoURI = $"https://teams.microsoft.com/l/stage/{TeamsInternalAppID}/0?context={this.UrlEncoder("Learn Together - Developing apps for Microsoft Teams", "https://www.youtube.com/embed/xxkCJKpU3vA", "https://www.youtube.com/watch?v=xxkCJKpU3vA")}";
                var video = new AdaptiveImageWithLongUrl
                {
                    LongUrl = posterLink,
                    Size = AdaptiveImageSize.Stretch,
                    HorizontalAlignment=AdaptiveHorizontalAlignment.Center,
                    SelectAction = new AdaptiveOpenUrlAction()
                    {

                        Url = new Uri(videoURI),

                    },


                };
                // Image enlarge support for Teams web/desktop client.
                video.AdditionalProperties.Add("msteams", new { AllowExpand = true });
                video.SelectAction.AdditionalProperties.Add("msteams", new { AllowExpand = true });

                card.Body.Add(video);

            }

            if (!string.IsNullOrWhiteSpace(summary))
            {
                card.Body.Add(new AdaptiveTextBlock()
                {
                    Text = summary,
                    Wrap = true,
                });
            }

            if (!string.IsNullOrWhiteSpace(author))
            {
                card.Body.Add(new AdaptiveTextBlock()
                {
                    Text = author,
                    Size = AdaptiveTextSize.Small,
                    Weight = AdaptiveTextWeight.Lighter,
                    Wrap = true,
                });
            }

            if (!string.IsNullOrWhiteSpace(buttonTitle)
                && !string.IsNullOrWhiteSpace(buttonUrl))
            {
                card.Actions.Add(new AdaptiveOpenUrlAction()
                {
                    Title = buttonTitle,
                    Url = new Uri(buttonUrl, UriKind.RelativeOrAbsolute),
                });
            }

            if(template != "departmentVideo" || template !="Deafult")
            {
                var columnSet = new AdaptiveColumnSet()
                {
                    Columns = new List<AdaptiveColumn>() {
                    new AdaptiveColumn()
                    {
                        Width = AdaptiveColumnWidth.Stretch,

                        Items = new List<AdaptiveElement>()
                        {
                            new AdaptiveImage()
                            {
                                Url = new Uri(Constants.BaseUrl + "/image/Customs.png", UriKind.RelativeOrAbsolute),
                                Size = AdaptiveImageSize.Stretch,
                            },
                        },
                    },
                    new AdaptiveColumn()
                    {
                        Width = AdaptiveColumnWidth.Stretch,

                        Items = new List<AdaptiveElement>()
                        {
                            new AdaptiveImage()
                            {
                                Url = new Uri(Constants.BaseUrl + "/image/UAE.png", UriKind.RelativeOrAbsolute),
                                Size = AdaptiveImageSize.Stretch,
                            },
                        },
                    },
                    },
                }; 
                card.Body.Add(columnSet);
            }

            // Full width Adaptive card.
            card.AdditionalProperties.Add("msteams", new { width = "full" });
            return card;
        }
    }
}