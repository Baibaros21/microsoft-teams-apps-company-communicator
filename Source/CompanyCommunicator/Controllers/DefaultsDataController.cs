using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Localization;
using Microsoft.Extensions.Configuration;
using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
using Microsoft.Teams.Apps.CompanyCommunicator.Common;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Resources;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Blob;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Teams;
using Microsoft.Teams.Apps.CompanyCommunicator.DraftNotificationPreview;
using Microsoft.Teams.Apps.CompanyCommunicator.Models;
using Microsoft.Teams.Apps.CompanyCommunicator.Repositories.Extensions;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Blob;


namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    /// <summary>
    /// Controller for the draft notification data.
    /// </summary>
    [Route("api/defaultdata")]
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class DefaultsDataController : ControllerBase
    {
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly ITeamDataRepository teamDataRepository;
        private readonly IDraftNotificationPreviewService draftNotificationPreviewService;
        private readonly IGroupsService groupsService;
        private readonly IAppSettingsService appSettingsService;
        private readonly IStringLocalizer<Strings> localizer;
        private readonly IBlobStorageProvider blobStorageProvider;
        private readonly string DEFAULT_LOGO_BLOB_NAME = "DEFAULT_LOGO";
        private readonly string DEFAULT_BANNER_BLOB_NAME = "DEFAULT_BANNER";
        private readonly string DEFAULT_HEADER_LOGO = "DEFAULT_HEADER_LOGO";

        /// <summary>
        /// Gets the IConfiguration instance.
        /// </summary>
        public IConfiguration Configuration { get; }

        public DefaultsDataController(
            INotificationDataRepository notificationDataRepository,
            ITeamDataRepository teamDataRepository,
            IDraftNotificationPreviewService draftNotificationPreviewService,
            IAppSettingsService appSettingsService,
            IStringLocalizer<Strings> localizer,
            IGroupsService groupsService,
            IBlobStorageProvider blobStorageProvider,
            IConfiguration configuration)
        {
            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentNullException(nameof(notificationDataRepository));
            this.teamDataRepository = teamDataRepository ?? throw new ArgumentNullException(nameof(teamDataRepository));
            this.draftNotificationPreviewService = draftNotificationPreviewService ?? throw new ArgumentNullException(nameof(draftNotificationPreviewService));
            this.localizer = localizer ?? throw new ArgumentNullException(nameof(localizer));
            this.groupsService = groupsService ?? throw new ArgumentNullException(nameof(groupsService));
            this.appSettingsService = appSettingsService ?? throw new ArgumentNullException(nameof(appSettingsService));
            this.blobStorageProvider = blobStorageProvider ?? throw new ArgumentException(nameof(blobStorageProvider));
            this.Configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));

        }

        [HttpPut]
        public async Task<IActionResult> UpdateDefaultsDataAsync([FromBody] DefaultsData defaults)
        {

            if (!string.IsNullOrEmpty(defaults.HeaderLogoLink) && defaults.HeaderLogoLink.StartsWith(Constants.ImageBase64Format))
            {
                await this.notificationDataRepository.SaveImageAsync(this.DEFAULT_HEADER_LOGO, defaults.HeaderLogoLink);
            }

            if (!string.IsNullOrEmpty(defaults.LogoLink) && defaults.LogoLink.StartsWith(Constants.ImageBase64Format))
            {
                await this.notificationDataRepository.SaveImageAsync(this.DEFAULT_LOGO_BLOB_NAME, defaults.LogoLink);

            }
            if (!string.IsNullOrEmpty(defaults.BannerLink) && defaults.LogoLink.StartsWith(Constants.ImageBase64Format))
            {
                await this.notificationDataRepository.SaveImageAsync(this.DEFAULT_BANNER_BLOB_NAME, defaults.BannerLink);

            }
            return this.Ok();
        }

        [HttpGet]

        public async Task<ActionResult<DefaultsData>> GetDefaultDataAsync()
        {

            var logoLink = "data:image/jpeg;base64," + await this.blobStorageProvider.DownloadBase64ImageAsync(this.DEFAULT_LOGO_BLOB_NAME);

            var bannerLink = "data:image/jpeg;base64," + await this.blobStorageProvider.DownloadBase64ImageAsync(this.DEFAULT_BANNER_BLOB_NAME);

            var headerLogoLink = "data:image/jpeg;base64," + await this.blobStorageProvider.DownloadBase64ImageAsync(this.DEFAULT_HEADER_LOGO);
            var result = new DefaultsData
            {
                LogoFileName = this.DEFAULT_LOGO_BLOB_NAME,
                LogoLink = logoLink,
                BannerFileName = this.DEFAULT_BANNER_BLOB_NAME,
                BannerLink = bannerLink,
                HeaderLogoLink = headerLogoLink,

            };



            return this.Ok(result);
        }

    }
}