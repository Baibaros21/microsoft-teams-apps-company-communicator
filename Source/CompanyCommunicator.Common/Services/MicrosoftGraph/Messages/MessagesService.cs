using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Configuration;
using Newtonsoft.Json.Linq;

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph.Messages
{
    /// <summary>
    /// MessageService
    /// </summary>

    internal class MessagesService : IMessageService
    {

        private readonly IGraphServiceClient graphServiceClient;
        private readonly IAppConfiguration appConfiguration;

        /// <summary>
        /// Initializes a new instance of the <see cref="MessagesService"/> class.
        /// </summary>
        /// <param name="graphServiceClient"></param>
        /// <param name="appConfiguration"></param>
        internal MessagesService(
            IGraphServiceClient graphServiceClient,
            IAppConfiguration appConfiguration)
        {
            this.graphServiceClient = graphServiceClient ?? throw new ArgumentNullException(nameof(graphServiceClient));
            this.appConfiguration = appConfiguration ?? throw new ArgumentNullException(nameof(appConfiguration));
        }

        public Task<Message> GetMessageAsync(string messageId)
        {
            
            throw new NotImplementedException();
        }
    }
}
