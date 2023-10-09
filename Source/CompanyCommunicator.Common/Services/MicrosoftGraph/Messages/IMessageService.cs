namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph.Messages
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net.Http;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Configuration;
    using Newtonsoft.Json.Linq;


    /// <summary>
    /// Get the Messages data
    /// </summary>
    public interface IMessageService
    {



        /// <summary>
        /// get message by id.
        /// </summary>
        /// <param name="messageId">the message id.</param>
        /// <returns>user data.</returns>
        public  Task<Message> GetMessageAsync(string messageId);
    }
}
