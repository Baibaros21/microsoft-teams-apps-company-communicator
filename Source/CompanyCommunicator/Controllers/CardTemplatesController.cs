namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net.Http;
    using Newtonsoft.Json.Linq;
    using System.Reactive;
    using System.Security.Claims;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Azure.WebJobs.Extensions.DurableTask;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Graph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Blob;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;
    using Newtonsoft.Json;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;

    /// <summary>
    /// Controller for the sent notification data.
    /// </summary>
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    [Route("api/cardtemplates")]
    public class CardTemplatesController : ControllerBase
    {
        private readonly IBlobStorageProvider storageProvider;


        /// <summary>
        /// Initializes a new instance of the <see cref="CardTemplatesController"/> class.
        /// </summary>
        public CardTemplatesController(
   IBlobStorageProvider storageProvider)
        {
            this.storageProvider = storageProvider ?? throw new ArgumentNullException(nameof(storageProvider));
        }


        [HttpGet]
        public async Task<ActionResult<IEnumerable<TemplateData>>> GetAllCardTemplates()
        {
            var cardTemplates = await this.storageProvider.DownloadAllCardTemplatesAsync();

            return this.Ok(cardTemplates);
        }

        [HttpGet("{id}")]

        public async Task<ActionResult<TemplateData>> GetCardTemplate(string id)
        {
            if (id == null)
            {
                throw new ArgumentNullException(nameof(id));
            }
            var result = await this.storageProvider.DownloadCardTemplateAsync(id);

            return this.Ok(result);
        }

        [HttpPut]
        public async Task<ActionResult> UpdateCardTemplate([FromBody] TemplateData cardTemplate)
        {
            if (cardTemplate == null)
            {
                throw new ArgumentNullException(nameof(cardTemplate));
            }

            await this.storageProvider.UploadCardTemplateAsync(cardTemplate.Template, cardTemplate.Card);
            return this.Ok();
        }

    }
}