#region using
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
#endregion

namespace Client.API.MicrosoftGraph.ApplicationPermission.Framework
{
    /// <summary>
    /// Base class for all services
    /// </summary>
    public class MsGraphService
    {
        protected readonly AppSettings _appSettings;
        private readonly IMemoryCache _cache;

        public MsGraphService(IOptions<AppSettings> options, IMemoryCache cache)
        {
            _appSettings = options.Value;
            _cache = cache;
        }

        #region MS Graph API
        /// <summary>
        /// Retrieve an access token to call MS Graph API
        /// </summary>
        /// <see cref="https://developer.microsoft.com/en-us/graph/docs/concepts/auth_v2_service"/>
        /// <seealso cref="https://stackoverflow.com/questions/37151346/authorization-identitynotfound-error-while-accessing-graph-api"/>
        /// <returns>access token</returns>
        private async Task<string> GetAccessToken()
        {
            // OAuth 2 Client Credential Flow Grant: https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-protocols-oauth-client-creds
            // Admin consent: https://login.microsoftonline.com/common/adminconsent?client_id=279e9c2f-0a8b-4e57-8d39-c40052edd0a7&state=12345&redirect_uri=https://localhost:44382/adminconsent

            var cacheKey = $"accessToken_{_appSettings.TenantId}";
            var accessToken = _cache.Get<string>(cacheKey);
            if (!string.IsNullOrWhiteSpace(accessToken))
            {
                return accessToken;
            }

            var token = await RequestAccessToken();
            if (token == null) throw new NullReferenceException("Access token cannot be NULL!");

            // Access token will be stocked in cache for one hour
            _cache.Set(cacheKey, token.access_token, new MemoryCacheEntryOptions
            {
                AbsoluteExpiration = DateTime.Now.AddSeconds(token.expires_in - 10)
            });

            return token.access_token;
        }

        /// <summary>
        /// Make a request to AAD v2 endpoint to obtain a new access token
        /// </summary>
        /// <returns></returns>
        private async Task<TokenResult> RequestAccessToken()
        {
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Accept", "application/json");

                var values = new List<KeyValuePair<string, string>>();
                values.Add(new KeyValuePair<string, string>("client_id", _appSettings.ClientId));
                values.Add(new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/.default"));
                values.Add(new KeyValuePair<string, string>("client_secret", _appSettings.ClientSecret));
                values.Add(new KeyValuePair<string, string>("grant_type", "client_credentials"));

                using (var content = new FormUrlEncodedContent(values))
                {
                    var requestUrl = _appSettings.AADTokenEndPointv2.Replace("common", _appSettings.TenantId);
                    using (var response = await client.PostAsync(requestUrl, content))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            return JsonConvert.DeserializeObject<TokenResult>(await response.Content.ReadAsStringAsync());
                        }
                        else
                        {
                            return null;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Return a authenticated client
        /// </summary>
        /// <returns></returns>
        public GraphServiceClient GetAuthenticatedClient()
        {
            return new GraphServiceClient(new DelegateAuthenticationProvider(
                async request =>
                {
                    // Request an access token
                    var accessToken = await GetAccessToken();

                    // Append the access token to the request
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                }));
        }

        public async Task<T> GetSingleResponse<T>(HttpRequestMessage request)
        {
            var client = GetAuthenticatedClient();

            await client.AuthenticationProvider.AuthenticateRequestAsync(request);

            var response = await client.HttpProvider.SendAsync(request);
            if (response.IsSuccessStatusCode)
            {
                var content = await response.Content.ReadAsStringAsync();
                return client.HttpProvider.Serializer.DeserializeObject<T>(content);
            }

            return default(T);
        }

        public async Task<T> GetListResponse<T>(HttpRequestMessage request)
        {
            var client = GetAuthenticatedClient();

            await client.AuthenticationProvider.AuthenticateRequestAsync(request);

            var response = await client.HttpProvider.SendAsync(request);
            if (response.IsSuccessStatusCode)
            {
                var content = JToken.Parse(await response.Content.ReadAsStringAsync())["value"].ToString();
                return client.HttpProvider.Serializer.DeserializeObject<T>(content);
            }

            return default(T);
        }
        #endregion

        #region Rest Sharp
        /// <summary>
        /// Create rest client
        /// </summary>
        /// <param name="endpoint">endpoint to request</param>
        /// <returns></returns>
        private IRestClient CreateClient(string endpoint)
        {
            return new RestClient(endpoint);
        }

        /// <summary>
        /// Create rest request with all headers required
        /// </summary>
        /// <param name="method"></param>
        /// <param name="payload"></param>
        /// <param name="headers"></param>
        /// <returns></returns>
        private IRestRequest CreateRequest(Method method, object payload, Dictionary<string, string> headers)
        {
            var request = new RestRequest(method);
            request.RequestFormat = DataFormat.Json;

            request.AddHeader("cache-control", "no-cache");
            request.AddHeader("Content-Type", "application/json");
            foreach (var header in headers)
            {
                request.AddHeader(header.Key, header.Value);
            }

            if (method == Method.GET || method == Method.DELETE) return request;

            request.AddBody(payload);
            return request;
        }

        /// <summary>
        /// Return response from api
        /// </summary>
        /// <param name="endpoint"></param>
        /// <param name="method"></param>
        /// <param name="headers"></param>
        /// <param name="payload"></param>
        /// <returns></returns>
        protected async Task<IRestResponse> GetResponse(string endpoint, Method method, Dictionary<string, string> headers, object payload = null)
        {
            var client = CreateClient(endpoint);
            var request = CreateRequest(method, payload, headers);

            // Return Rest response
            return await client.ExecuteTaskAsync(request);
        }

        /// <summary>
        /// Create api result
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="restResponse"></param>
        /// <returns></returns>
        protected T CreateApiResult<T>(IRestResponse restResponse)
        {
            if (restResponse.IsSuccessful)
            {
                return JsonConvert.DeserializeObject<T>(restResponse.Content);
            }

            return default(T);
        }

        #endregion
    }

    public class TokenResult
    {
        public string access_token { get; set; }
        public int expires_in { get; set; }
        public string token_type { get; set; }
    }
}
