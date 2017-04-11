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
        /// activate a solution which has already been uploaded to a document library (or Solutions library)
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="name">name of the solution</param>
        /// <param name="guid">unique guid for the solution</param>
        /// <param name="majorVersion"></param>
        /// <param name="minorVersion"></param>
        /// <param name="wspFileRelativePath">location of wsp file relative to the site eg /_catalogs/Solutions/abc.wsp</param>
        public void ActivateSolution(string siteUrl, string name, Guid guid, int majorVersion, int minorVersion, string wspFileRelativePath)
        {
            using (var context = GetContext(siteUrl))
            {
                
                _logger.LogVerbose(string.Format("Activating solution '{0}'...", name));

                if (majorVersion == 0) majorVersion = 1;

                DesignPackageInfo info = new DesignPackageInfo()
                {
                    PackageGuid = guid,
                    MajorVersion = majorVersion,
                    MinorVersion = minorVersion,
                    PackageName = name
                };
                
                DesignPackage.Install(context, context.Site, info, wspFileRelativePath);
                context.ExecuteQuery();
            }
        }




    }
}
