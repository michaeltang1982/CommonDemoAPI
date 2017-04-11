using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using SP = Microsoft.SharePoint.Client;
using Microsoft.Online.SharePoint.TenantAdministration;

namespace Sierra.SharePoint.Library.CSOM
{
    public class SPOnlineUtility
    {

        public static void DisableDenyAddAndCustomizePages(SP.ClientContext context, string siteUrl)
        {
            
            var tenant = new Tenant(context);
            var siteProperties = tenant.GetSitePropertiesByUrl(siteUrl, true);
            context.Load(siteProperties);
            context.ExecuteQuery();

            siteProperties.DenyAddAndCustomizePages = DenyAddAndCustomizePagesStatus.Disabled;
            var result = siteProperties.Update();
            context.Load(result);
            context.ExecuteQuery();
            while (!result.IsComplete)
            {
                Thread.Sleep(result.PollingInterval);
                context.Load(result);
                context.ExecuteQuery();
            }
            
        }
    }
}
