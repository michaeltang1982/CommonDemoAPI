using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Security;
using SP = Microsoft.SharePoint.Client;
using Sierra.NET.Core;

namespace Sierra.SharePoint.Library.CSOM
{
    public partial class SPClientUtility
    {
        private ILogger _logger;
        private System.Net.ICredentials _credentials;

        /// <summary>
        /// package up credentials for use in connecting to SharePoint
        /// </summary>
        /// <param name="userName">user name to use</param>
        /// <param name="password">(optional) password</param>
        /// <param name="isSPOnline">is this SP Online or on-premise?</param>
        /// <returns>ICredentials object</returns>
        public static System.Net.ICredentials GetCredentials(string userName, string password, bool isSPOnline)
        {

            if (isSPOnline)
                return new Microsoft.SharePoint.Client.SharePointOnlineCredentials(userName, ConvertToSecureString(password));
            else
            {
                //on-premise credentials
                if (string.IsNullOrEmpty(userName))
                    return System.Net.CredentialCache.DefaultCredentials;
                else
                {
                    return new System.Net.NetworkCredential(userName, ConvertToSecureString(password));

                }
            }


        }

        /// <summary>
        /// create SP Client Utility object
        /// </summary>
        /// <param name="logger"></param>
        /// <param name="credentials"></param>
        public SPClientUtility(ILogger logger, System.Net.ICredentials credentials)
        {
            _logger = logger;
            _credentials = credentials;
        }

        


        public string GetSiteUrl(string parentUrl, string relativeWebUrl)
        {
            return parentUrl.TrimEnd('/') + "/" + relativeWebUrl.TrimStart('/');
        }


        public string GetListFolderUrl(string listRelativePath, string folderRelativePath)
        {
            return "/" + listRelativePath.Trim('/') + "/" + folderRelativePath.Trim('/');
        }

        public SP.ClientContext GetContext(string url)
        {
            SP.ClientContext context = null;

            try
            {
                context = new SP.ClientContext(url);
            }
            catch (ArgumentException)
            {
                throw new Exception(string.Format("GetContext: No site found with the following url: " + url));
            }
            context.Credentials = _credentials;
            return context;
        }


        /// <summary>
        /// Combines a path and a relative path.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="relative"></param>
        /// <returns></returns>
        public static string CombinePath(string path, string relative)
        {
            if (relative == null)
                relative = String.Empty;

            if (path == null)
                path = String.Empty;

            if (relative.Length == 0 && path.Length == 0)
                return String.Empty;

            if (relative.Length == 0)
                return path;

            if (path.Length == 0)
                return relative;

            path = path.Replace('\\', '/');
            relative = relative.Replace('\\', '/');

            return path.TrimEnd('/') + '/' + relative.TrimStart('/');
        }


        private static SecureString ConvertToSecureString(string input)
        {
            var secureString = new SecureString();
            if (input.Length > 0)
            {
                char[] charArray = input.ToCharArray();
                foreach (var c in charArray) secureString.AppendChar(c);
            }
            return secureString;
        }
    }
}
