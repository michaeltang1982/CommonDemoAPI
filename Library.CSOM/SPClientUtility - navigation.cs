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
        public class SiteNavigationSettings
        {
            public bool inheritGlobal;
            public bool globalIncludeSubSites;
            public bool globalIncludePages;
            public bool inheritCurrent;
            public bool showSiblings;
            public bool currentIncludeSubSites;
            public bool currentIncludePages;
        }

        public void SetSiteNavigation(string url, SiteNavigationSettings navSettings)
        {
            using (var context = GetContext(url))
            {
                SP.Web web = context.Web;
                 
                _logger.LogVerbose("Setting navigation...");

                 
                //ClientPortalNavigation nav = new ClientPortalNavigation(web);

                //nav.GlobalIncludePages = navSettings.globalIncludePages;
                //nav.GlobalIncludeSubSites = navSettings.globalIncludeSubSites;

                //nav.CurrentIncludePages = navSettings.currentIncludePages;
                //nav.CurrentIncludeSubSites = navSettings.currentIncludeSubSites;

                //nav.InheritGlobalNavigation = navSettings.inheritGlobal;
                //nav.InheritCurrentNavigation = navSettings.inheritCurrent;


                //nav.SaveChanges();


                context.Load(web, w => w.AllProperties);
                context.ExecuteQuery();
                web.AllProperties["__GlobalNavigationIncludeTypes"] = GetNavTypesValue(navSettings.globalIncludeSubSites, navSettings.globalIncludePages);
                web.AllProperties["__CurrentNavigationIncludeTypes"] = GetNavTypesValue(navSettings.currentIncludeSubSites, navSettings.currentIncludePages);

                web.Update();
                context.ExecuteQuery();

                var webNavigationSettings = new WebNavigationSettings(context, web);

                webNavigationSettings.GlobalNavigation.Source = (navSettings.inheritGlobal) ? StandardNavigationSource.InheritFromParentWeb : StandardNavigationSource.PortalProvider;


                webNavigationSettings.CurrentNavigation.Source = (navSettings.inheritCurrent) ? StandardNavigationSource.InheritFromParentWeb : StandardNavigationSource.PortalProvider;

                TaxonomySession taxSession = TaxonomySession.GetTaxonomySession(context);
                webNavigationSettings.Update(taxSession);
                web.Update();
                context.ExecuteQuery();
            }
        }


        private int GetNavTypesValue(bool includeSites, bool includePages)
        {
            //0 = don't show pages or subsites
            //1 = show subsites only
            //2 = show pages only
            //3 = show subsites and pages 
            int s = includeSites ? 1 : 0;
            int p = includePages ? 2 : 0;
            return s + p;

        }
        public void CreateNavigationNode(string siteUrl, string relativeWebUrl, string title, string webTemplate, bool deleteIfExists, bool breakRoleInheritance, bool copyParentRoles)
        {
            

            using (var context = GetContext(siteUrl))
            {
                _logger.LogVerbose("Creating site..." + title);
                context.Load(context.Web);

                SP.NavigationNodeCollection navNodes = context.Web.Navigation.QuickLaunch;


                IEnumerable<SP.NavigationNode> twitterNode = context.LoadQuery(navNodes.Where(n => n.Title == "Twitter"));


                
                try
                {
                    context.ExecuteQuery();
                }
                catch (Exception)
                {
                    //Handle exception
                }

                //Create a new node

                SP.NavigationNodeCreationInformation newNavNode = new SP.NavigationNodeCreationInformation();

                newNavNode.Title = "Google";

                newNavNode.Url = "https://www.google.com"; //URL must always start with http/https

                newNavNode.PreviousNode = twitterNode.FirstOrDefault();

                navNodes.Add(newNavNode);

                try
                {
                    context.ExecuteQuery();
                }
                catch (Exception)
                {
                    //Handle exception
                }

            }
        }


        private void UpdateNavigationNode(string siteUrl)
        {
            using (var context = GetContext(siteUrl))
            {

                context.Load(context.Web);

                //Fetching website's Left Navigation node collection

                SP.NavigationNodeCollection qlNavNodeColl = context.Web.Navigation.QuickLaunch;

                //Fetching node which needs to be updated

                IEnumerable<SP.NavigationNode> googleNode = context.LoadQuery(qlNavNodeColl.Where(n => n.Title == "Google"));

                try
                {
                    context.ExecuteQuery();
                }
                catch (Exception)
                {
                    //Handle exception
                }

                if (googleNode.Count() == 1)
                {
                    SP.NavigationNode gNode = googleNode.FirstOrDefault();

                    gNode.Url = "https://www.google.co.in";

                    gNode.Update();

                    try
                    {
                        context.ExecuteQuery();
                    }
                    catch (Exception)
                    {
                        //Handle exception
                    }
                }
            }
        }

        private void DeleteNavigationNode(string siteUrl)
        {
            using (var context = GetContext(siteUrl))
            {

                context.Load(context.Web);

                //Fetching website's Left Navigation node collection

                SP.NavigationNodeCollection qlNavNodeColl = context.Web.Navigation.QuickLaunch;

                //Fetching node which needs to be deleted

                IEnumerable<SP.NavigationNode> googleNode = context.LoadQuery(qlNavNodeColl.Where(n => n.Title == "Google"));

                try
                {
                    context.ExecuteQuery();
                }
                catch (Exception)
                {
                    //Handle exception
                }

                if (googleNode.Count() == 1)
                {
                    googleNode.FirstOrDefault().DeleteObject();

                    try
                    {
                        context.ExecuteQuery();
                    }
                    catch (Exception)
                    {
                        //Handle exception
                    }
                }
            }
        }
        
    }
}
