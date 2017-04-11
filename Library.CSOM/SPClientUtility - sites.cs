using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client;


namespace Sierra.SharePoint.Library.CSOM
{
    public partial class SPClientUtility
    {

        /// <summary>
        /// get all the urls of the subsites
        /// </summary>
        /// <param name="siteUrl">parent site url</param>
        /// <returns></returns>
        public List<string> GetSubSiteUrls(string siteUrl)
        {
            
            using (var context = GetContext(siteUrl))
            {

                SP.Web web = context.Web;
                List<string> urls = new List<string>();
                GetAllSubSiteUrls(context, web, urls);

                return urls;

            }
            

        }


        /// <summary>
        /// get all the urls of the subsites
        /// </summary>
        /// <param name="siteUrl">parent site url</param>
        /// <returns></returns>
        public List<string> GetSiteProperties(string siteUrl, params string[] fieldNames)
        {

            using (var context = GetContext(siteUrl))
            {

                SP.Web web = context.Web;
                List<string> urls = new List<string>();
                
                foreach (var fieldName in fieldNames)
                {
                    //context.Load(web, w => w[fieldName]));
                }
                context.ExecuteQuery();
                         

                return urls;

            }


        }

        /// <summary>
        /// get list of features activate at SPSite level
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <returns></returns>
        public List<string> GetActivatedFeaturesForSite(string siteUrl)
        {
            _logger.LogVerbose("Getting site collection features for url: " + siteUrl);

            List<string> siteFeatures = new List<string>();

            using (var context = GetContext(siteUrl))
            {
                FeatureCollection features = context.Site.Features;
                context.Load(features, f=>f.Include(m=>m.DefinitionId));
                context.ExecuteQuery();

                foreach(var feature in features)  siteFeatures.Add(feature.DefinitionId.ToString());
            }

            return siteFeatures;
        }

        public void CreateWeb(string parentUrl, string relativeWebUrl, string title, string webTemplate, bool deleteIfExists, bool breakRoleInheritance, bool copyParentRoles)
        {
            relativeWebUrl = relativeWebUrl.TrimStart('/');

            if (!string.IsNullOrEmpty(relativeWebUrl))
            {
                if (deleteIfExists)
                {
                    DeleteSite(GetSiteUrl(parentUrl, relativeWebUrl));
                }

                using (var context = GetContext(parentUrl))
                {
                    context.Load(context.Web.Webs, w => w.Include(x => x.ServerRelativeUrl));
                    context.ExecuteQuery();

                    if (!context.Web.Webs.Any(x => x.ServerRelativeUrl.EndsWith("/" + relativeWebUrl, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        _logger.LogVerbose("Creating site..." + title);
                        SP.WebCreationInformation creation = new SP.WebCreationInformation();
                        creation.Url = relativeWebUrl;
                        creation.Title = title;
                        creation.WebTemplate = webTemplate;
                        creation.UseSamePermissionsAsParentSite = !breakRoleInheritance;

                        SP.Web newWeb = context.Web.Webs.Add(creation);

                        context.ExecuteQuery();

                    }
                    else
                        _logger.LogVerbose("Site with following url already exists: " + relativeWebUrl);
                }
            }
            else
                _logger.LogVerbose("This is the root site for the collection");
        }

        public void UpdateSite(string url, bool breakRoleInheritance, string homePageUrl)
        {
            using (var context = GetContext(url))
            {
                SP.Web web = context.Web;
                
                _logger.LogVerbose("Updating site at url: " + url);
                    
                if (breakRoleInheritance)
                {
                    _logger.LogVerbose("Breaking inheritance...");
                    web.BreakRoleInheritance(false, true);
                    context.ExecuteQuery();
                }


            }
        }


        


        public void DeleteSite(string fullSiteUrl)
        {
            try
            {
                using (var context = GetContext(fullSiteUrl))
                {
                    
                    SP.Web web = context.Web;
                    List<string> urls = new List<string>();
                    GetAllSubSiteUrls(context, web, urls);

                    foreach(string url in urls)
                    {
                        _logger.LogVerbose("Deleting site: " + url);
                        DeleteSiteAtUrl(url);
                    }
                    
                }
            }
            catch (SP.ClientRequestException ex)
            {
                if (!ex.Message.ToLower().Contains("there is no web"))
                    throw;
            }

        }


        private void DeleteSite(SP.ClientContext context, SP.Web web)
        {
            var subwebs = web.Webs;
            context.Load(web);            
            context.Load(subwebs);
            context.ExecuteQuery();
            
            if (subwebs.Count > 0)
            {

                foreach (var subweb in subwebs)
                {
                    DeleteSite(context, subweb);
                }

            }

            _logger.LogVerbose("Deleting site: " + web.Url);
            web.DeleteObject();
            context.ExecuteQuery();
        }

        private void DeleteSiteAtUrl(string siteUrl)
        {
            using (var context = GetContext(siteUrl))
            {
                SP.Web web = context.Web;
                web.DeleteObject();
                context.ExecuteQuery();
            }
   
        }

        private void GetAllSubSiteUrls(SP.ClientContext context, SP.Web web, List<string> urls)
        {
            var subwebs = web.Webs;
            context.Load(web);
            context.Load(subwebs, s=>s.Include(k=>k.Url));
            context.ExecuteQuery();

            if (subwebs.Count > 0)
            {

                foreach (var subweb in subwebs)
                {
                    GetAllSubSiteUrls(context, subweb, urls);
                }

            }

            urls.Add(web.Url);
        }

        public void ActivateFeature(string siteUrl, Guid featureId, bool isSitefeature)
        {
            _logger.LogVerbose("Activating feature...");
            using (var context = GetContext(siteUrl))
            {
                SP.FeatureCollection features = null;
                if (isSitefeature)
                { 
                    features = context.Site.Features;
                    
                }
                else
                {
                    features = context.Web.Features;
                }

                features.Add(featureId, true, SP.FeatureDefinitionScope.None);

                context.ExecuteQuery();
            } 
        }

        public void DeactivateFeature(string siteUrl, Guid featureId, bool isSitefeature)
        {
            _logger.LogVerbose("Deactivating feature...");
            using (var context = GetContext(siteUrl))
            {
                SP.FeatureCollection features = null;

                if (isSitefeature)
                {
                    SP.Site site = context.Site;
                    context.Load(site, w => w.Features);
                    context.ExecuteQuery();
                    features = site.Features;
                }
                else
                {
                    SP.Web web = context.Web;
                    context.Load(web, w => w.Features);
                    context.ExecuteQuery();
                    features = web.Features;
                }
                


                if (features.Where(w => w.DefinitionId == featureId).Count() > 0)
                {
                    
                    features.Remove(featureId, true);
                    context.ExecuteQuery();
                }
                else
                {
                    _logger.LogVerbose("Feature already de-activated");
                }
            }
        }


        
        
    }
}
