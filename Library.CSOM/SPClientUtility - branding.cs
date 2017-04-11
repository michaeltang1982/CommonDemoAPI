using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client.Publishing.Navigation;


namespace Sierra.SharePoint.Library.CSOM
{
    public partial class SPClientUtility
    {
        

        

        /// <summary>
        /// set alternate css and site logo
        /// </summary>
        /// <param name="siteUrl">full url of the site</param>
        /// <param name="siteRelativeAlternateCssUrl">css url relative to the site collection  eg "/SubSite/SiteAssets/my.css" </param>
        /// <param name="siteRelativeSiteLogoUrl"></param>
        public void SetBranding(string siteUrl, string siteRelativeAlternateCssUrl, string siteRelativeSiteLogoUrl)
        {
            using (var context = GetContext(siteUrl))
            {
                SP.Web web = context.Web;

                _logger.LogVerbose("Branding: setting logo and alternate css...");

                // Set the properties accordingly
                // NOTE: these props need at least 16.0.3104.1200 of the client dlls
                if (!string.IsNullOrEmpty(siteRelativeAlternateCssUrl))
                    web.AlternateCssUrl = siteRelativeAlternateCssUrl; // "/Resources/SiteAssets/my.css";
                else
                    web.AlternateCssUrl = string.Empty; //clear this property

                if (!string.IsNullOrEmpty(siteRelativeSiteLogoUrl))
                    web.SiteLogoUrl = siteRelativeSiteLogoUrl; // "/Resources/SiteAssets/my.png";
                else
                    web.SiteLogoUrl = string.Empty; //clear it

                web.Update();
                web.Context.ExecuteQuery();
            }
        }




    }
}
