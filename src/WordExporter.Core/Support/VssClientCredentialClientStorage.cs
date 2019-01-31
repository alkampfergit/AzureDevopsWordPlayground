using Microsoft.VisualStudio.Services.Client;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.Common.TokenStorage;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace WordExporter.Core.Support
{
    /// <summary>
    /// https://stackoverflow.com/questions/47230981/vssconnection-to-vsts-always-prompts-for-credentials
    /// Class to alter the credential storage behavior to allow the token to be cached between sessions.
    /// </summary>
    /// <seealso cref="Microsoft.VisualStudio.Services.Common.IVssCredentialStorage" />
    public class VssClientCredentialCachingStorage : VssClientCredentialStorage
    {
        #region [Private]

        private const string __tokenExpirationKey = "ExpirationDateTime";
        private double _tokenLeaseInSeconds;

        #endregion [Private]

        /// <summary>
        /// The default token lease in seconds
        /// </summary>
        public const double DefaultTokenLeaseInSeconds = 86400;// one day

        /// <summary>
        /// Initializes a new instance of the <see cref="VssClientCredentialCachingStorage"/> class.
        /// </summary>
        /// <param name="storageKind">Kind of the storage.</param>
        /// <param name="storageNamespace">The storage namespace.</param>
        /// <param name="tokenLeaseInSeconds">The token lease in seconds.</param>
        public VssClientCredentialCachingStorage(string storageKind = "VssApp", string storageNamespace = "VisualStudio", double tokenLeaseInSeconds = DefaultTokenLeaseInSeconds)
            : base(storageKind, storageNamespace)
        {
            this._tokenLeaseInSeconds = tokenLeaseInSeconds;
        }

        /// <summary>
        /// Removes the token.
        /// </summary>
        /// <param name="serverUrl">The server URL.</param>
        /// <param name="token">The token.</param>
        public override void RemoveToken(Uri serverUrl, IssuedToken token)
        {
            this.RemoveToken(serverUrl, token, false);
        }

        /// <summary>
        /// Removes the token.
        /// </summary>
        /// <param name="serverUrl">The server URL.</param>
        /// <param name="token">The token.</param>
        /// <param name="force">if set to <c>true</c> force the removal of the token.</param>
        public void RemoveToken(Uri serverUrl, IssuedToken token, bool force)
        {
            //////////////////////////////////////////////////////////
            // Bypassing this allows the token to be stored in local
            // cache. Token is removed if lease is expired.

            if (force || token != null && this.IsTokenExpired(token))
                base.RemoveToken(serverUrl, token);

            //////////////////////////////////////////////////////////
        }

        /// <summary>
        /// Retrieves the token.
        /// </summary>
        /// <param name="serverUrl">The server URL.</param>
        /// <param name="credentialsType">Type of the credentials.</param>
        /// <returns>The <see cref="IssuedToken"/></returns>
        public override IssuedToken RetrieveToken(Uri serverUrl, VssCredentialsType credentialsType)
        {
            var token = base.RetrieveToken(serverUrl, credentialsType);

            if (token != null)
            {
                bool expireToken = this.IsTokenExpired(token);
                if (expireToken)
                {
                    base.RemoveToken(serverUrl, token);
                    token = null;
                }
                else
                {
                    // if retrieving the token before it is expired,
                    // refresh the lease period.
                    this.RefreshLeaseAndStoreToken(serverUrl, token);
                    token = base.RetrieveToken(serverUrl, credentialsType);
                }
            }

            return token;
        }

        /// <summary>
        /// Stores the token.
        /// </summary>
        /// <param name="serverUrl">The server URL.</param>
        /// <param name="token">The token.</param>
        public override void StoreToken(Uri serverUrl, IssuedToken token)
        {
            this.RefreshLeaseAndStoreToken(serverUrl, token);
        }

        /// <summary>
        /// Clears all tokens.
        /// </summary>
        /// <param name="url">The URL.</param>
        public void ClearAllTokens(Uri url = null)
        {
            IEnumerable<VssToken> tokens = this.TokenStorage.RetrieveAll(base.TokenKind).ToList();

            if (url != default(Uri))
                tokens = tokens.Where(t => StringComparer.InvariantCultureIgnoreCase.Compare(t.Resource, url.ToString().TrimEnd('/')) == 0);

            foreach (var token in tokens)
                this.TokenStorage.Remove(token);
        }

        private void RefreshLeaseAndStoreToken(Uri serverUrl, IssuedToken token)
        {
            if (token.Properties == null)
                token.Properties = new Dictionary<string, string>();

            token.Properties[__tokenExpirationKey] = JsonConvert.SerializeObject(this.GetNewExpirationDateTime());

            base.StoreToken(serverUrl, token);
        }

        private DateTime GetNewExpirationDateTime()
        {
            var now = DateTime.Now;

            // Ensure we don't overflow the max DateTime value
            var lease = Math.Min((DateTime.MaxValue - now.Add(TimeSpan.FromSeconds(1))).TotalSeconds, this._tokenLeaseInSeconds);

            // ensure we don't have negative leases
            lease = Math.Max(lease, 0);

            return now.AddSeconds(lease);
        }

        private bool IsTokenExpired(IssuedToken token)
        {
            bool expireToken = true;

            if (token != null && token.Properties.ContainsKey(__tokenExpirationKey))
            {
                string expirationDateTimeJson = token.Properties[__tokenExpirationKey];

                try
                {
                    DateTime expiration = JsonConvert.DeserializeObject<DateTime>(expirationDateTimeJson);

                    expireToken = DateTime.Compare(DateTime.Now, expiration) >= 0;
                }
                catch { }
            }

            return expireToken;
        }
    }
}
