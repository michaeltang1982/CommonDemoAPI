using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using Microsoft.SharePoint.Client.Publishing;


namespace Sierra.SharePoint.Library.CSOM
{
    public enum TypeOfPage
    {
        Unknown,
        FormPage,
        StandardPage,
        WikiPage,
        PublishingPage
    }

    /// <summary>
    /// Out of the box wiki page layouts enumeration
    /// </summary>
    public enum WikiPageLayout
    {
        OneColumn = 0,
        OneColumnSideBar = 1,
        TwoColumns = 2,
        TwoColumnsHeader = 3,
        TwoColumnsHeaderFooter = 4,
        ThreeColumns = 5,
        ThreeColumnsHeader = 6,
        ThreeColumnsHeaderFooter = 7
    }

    public partial class SPClientUtility
    {

        private const string WikiPage_OneColumn = @"<div class=""ExternalClassC1FD57BEDB8942DC99A06C02F9A98241""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td style=""width&#58;100%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">false,false,1</span></div>";
        private const string WikiPage_OneColumnSideBar = @"<div class=""ExternalClass47565ACDF7974263AA4A556DA974B687""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td style=""width&#58;66.6%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">false,false,2</span></div>";
        private const string WikiPage_TwoColumns = @"<div class=""ExternalClass3811C839E5984CCEA4C8CF738462AFD8""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td style=""width&#58;49.95%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;49.95%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">false,false,2</span></div>";
        private const string WikiPage_TwoColumnsHeader = @"<div class=""ExternalClass850251EB51394304A07A64A05C0BB0F1""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td colspan=""2""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr><tr style=""vertical-align&#58;top;""><td style=""width&#58;49.95%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;49.95%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">true,false,2</span></div>";
        private const string WikiPage_TwoColumnsHeaderFooter = @"<div class=""ExternalClass71C5527252AD45859FA774445D4909A2""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td colspan=""2""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr><tr style=""vertical-align&#58;top;""><td style=""width&#58;49.95%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;49.95%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr><tr style=""vertical-align&#58;top;""><td colspan=""2""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">true,true,2</span></div>";
        private const string WikiPage_ThreeColumns = @"<div class=""ExternalClass833D1FA704C94892A26C4069C3FE5FE9""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">false,false,3</span></div>";
        private const string WikiPage_ThreeColumnsHeader = @"<div class=""ExternalClassD1A150D6187F449B8A6C4BEA2D4913BB""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td colspan=""3""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr><tr style=""vertical-align&#58;top;""><td style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">true,false,3</span></div>";
        private const string WikiPage_ThreeColumnsHeaderFooter = @"<div class=""ExternalClass5849C2C61FEC44E9B249C60F7B0ACA38""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td colspan=""3""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr><tr style=""vertical-align&#58;top;""><td style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr><tr style=""vertical-align&#58;top;""><td colspan=""3""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">true,true,3</span></div>";


        /// <summary>
        /// create new wiki page
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="pagesListTitle"></param>
        /// <param name="pageFileName"></param>
        /// <param name="pageTitle"></param>
        /// <param name="pageType"></param>
        /// <param name="deleteIfExists"></param>
        /// <returns>server relative url of the page</returns>
        public string CreateWikiPage(string siteUrl, string pagesListTitle, string pageFileName, TypeOfPage pageType, string pageTitle, WikiPageLayout wikiPageLayout, bool deleteIfExists)
        {
            string pageUrl = string.Empty;

            string pageFileNameWithoutExtension = pageFileName.Replace(".aspx", "");
            if (string.IsNullOrEmpty(pageTitle)) pageTitle = pageFileNameWithoutExtension;
            pageFileName =  pageFileNameWithoutExtension + ".aspx";

            SP.TemplateFileType template = (SP.TemplateFileType)Enum.Parse(typeof(Microsoft.SharePoint.Client.TemplateFileType), pageType.ToString());

            _logger.LogVerbose(string.Format("Creation of wiki page '{0}' in library '{1}': ", pageFileName, pagesListTitle));

            using (var context = GetContext(siteUrl))
            {
                var web = context.Web;

                context.Load(web);
                context.Load(web.Lists);

                var pageLibrary = this.GetListByTitle(context, pagesListTitle);

                context.Load(pageLibrary.RootFolder, f => f.ServerRelativeUrl);
                context.ExecuteQuery();

                pageUrl = string.Format("/{0}/{1}/{2}", pageLibrary.RootFolder.ServerRelativeUrl, "", pageFileName).Replace("//", "/");

                _logger.LogVerbose("Finding existing page at: " + pageUrl);

                var currentPageFile = web.GetFileByServerRelativeUrl(pageUrl);
                
                context.Load(currentPageFile, f => f.Exists);
                context.ExecuteQuery();
                bool exists = currentPageFile.Exists;

                if (exists && deleteIfExists)
                {
                    _logger.LogVerbose("Deleting existing page...");
                    currentPageFile.DeleteObject();
                    context.ExecuteQuery();
                    exists = false;
                }

                if (!exists)
                {
                    _logger.LogVerbose("Creating page...");
                    var newpage = pageLibrary.RootFolder.Files.AddTemplateFile(pageUrl, template);

                    context.Load(newpage, f => f.ListItemAllFields);                    
                    SP.ListItem item = newpage.ListItemAllFields;
                    item["Title"] = pageTitle;
                    item.Update();
                    context.ExecuteQuery();

                    this.AddLayoutToWikiPage(web, wikiPageLayout, pageUrl);
                }
                else
                {
                    _logger.LogVerbose("Page already exists.");
                }
            }

            return pageUrl;


        }


        /// <summary>
        /// Applies a layout to a wiki page
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="layout">Wiki page layout to be applied</param>
        /// <param name="serverRelativePageUrl"></param>
        /// <exception cref="System.ArgumentException">Thrown when serverRelativePageUrl is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when serverRelativePageUrl is null</exception>
        public void AddLayoutToWikiPage(SP.Web web, WikiPageLayout layout, string serverRelativePageUrl)
        {
            if (string.IsNullOrEmpty(serverRelativePageUrl)) throw new ArgumentNullException("serverRelativePageUrl");

            _logger.LogVerbose("Applying wiki page layout...");

            string html = "";
            switch (layout)
            {
                case WikiPageLayout.OneColumn:
                    html = WikiPage_OneColumn;
                    break;
                case WikiPageLayout.OneColumnSideBar:
                    html = WikiPage_OneColumnSideBar;
                    break;
                case WikiPageLayout.TwoColumns:
                    html = WikiPage_TwoColumns;
                    break;
                case WikiPageLayout.TwoColumnsHeader:
                    html = WikiPage_TwoColumnsHeader;
                    break;
                case WikiPageLayout.TwoColumnsHeaderFooter:
                    html = WikiPage_TwoColumnsHeaderFooter;
                    break;
                case WikiPageLayout.ThreeColumns:
                    html = WikiPage_ThreeColumns;
                    break;
                case WikiPageLayout.ThreeColumnsHeader:
                    html = WikiPage_ThreeColumnsHeader;
                    break;
                case WikiPageLayout.ThreeColumnsHeaderFooter:
                    html = WikiPage_ThreeColumnsHeaderFooter;
                    break;
                default:
                    break;
            }

            this.AddHtmlToWikiPage(web, serverRelativePageUrl, html);
        }

        /// <summary>
        /// add piece of html to an existing wiki page
        /// </summary>
        public void AddHtmlToWikiPage(SP.Web web, string serverRelativePageUrl, string html)
        {
            if (string.IsNullOrEmpty(serverRelativePageUrl)) throw new ArgumentNullException("serverRelativePageUrl");
            if (string.IsNullOrEmpty(html)) throw new ArgumentNullException("html");

            SP.File file = web.GetFileByServerRelativeUrl(serverRelativePageUrl);

            web.Context.Load(file, f => f.ListItemAllFields);
            web.Context.ExecuteQuery();

            SP.ListItem item = file.ListItemAllFields;

            item["WikiField"] = html;
            item.Update();
            web.Context.ExecuteQuery();
        }



        public void AddWebPartToWikiPage(SP.Web web, string serverRelativePageUrl, string webPartXml, int row, int col, bool addSpace)
        {
            if (string.IsNullOrEmpty(serverRelativePageUrl)) throw new ArgumentNullException("serverRelativePageUrl");
            if (string.IsNullOrEmpty(webPartXml)) throw new ArgumentNullException("webPartXml");

            SP.File webPartPage = web.GetFileByServerRelativeUrl(serverRelativePageUrl);

            if (webPartPage == null)
            {
                throw new Exception("No page found at: " + serverRelativePageUrl);
            }

            web.Context.Load(webPartPage);
            web.Context.Load(webPartPage.ListItemAllFields);
            web.Context.ExecuteQuery();

            string wikiField = (string)webPartPage.ListItemAllFields["WikiField"];

            LimitedWebPartManager limitedWebPartManager = webPartPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            WebPartDefinition oWebPartDefinition = limitedWebPartManager.ImportWebPart(webPartXml);
            WebPartDefinition wpdNew = limitedWebPartManager.AddWebPart(oWebPartDefinition.WebPart, "wpz", 0);
            web.Context.Load(wpdNew);
            web.Context.ExecuteQuery();

            //HTML structure in default team site home page (W16)
            //<div class="ExternalClass284FC748CB4242F6808DE69314A7C981">
            //  <div class="ExternalClass5B1565E02FCA4F22A89640AC10DB16F3">
            //    <table id="layoutsTable" style="width&#58;100%;">
            //      <tbody>
            //        <tr style="vertical-align&#58;top;">
            //          <td colspan="2">
            //            <div class="ms-rte-layoutszone-outer" style="width&#58;100%;">
            //              <div class="ms-rte-layoutszone-inner" style="word-wrap&#58;break-word;margin&#58;0px;border&#58;0px;">
            //                <div><span><span><div class="ms-rtestate-read ms-rte-wpbox"><div class="ms-rtestate-read 9ed0c0ac-54d0-4460-9f1c-7e98655b0847" id="div_9ed0c0ac-54d0-4460-9f1c-7e98655b0847"></div><div class="ms-rtestate-read" id="vid_9ed0c0ac-54d0-4460-9f1c-7e98655b0847" style="display&#58;none;"></div></div></span></span><p> </p></div>
            //                <div class="ms-rtestate-read ms-rte-wpbox">
            //                  <div class="ms-rtestate-read c7a1f9a9-4e27-4aa3-878b-c8c6c87961c0" id="div_c7a1f9a9-4e27-4aa3-878b-c8c6c87961c0"></div>
            //                  <div class="ms-rtestate-read" id="vid_c7a1f9a9-4e27-4aa3-878b-c8c6c87961c0" style="display&#58;none;"></div>
            //                </div>
            //              </div>
            //            </div>
            //          </td>
            //        </tr>
            //        <tr style="vertical-align&#58;top;">
            //          <td style="width&#58;49.95%;">
            //            <div class="ms-rte-layoutszone-outer" style="width&#58;100%;">
            //              <div class="ms-rte-layoutszone-inner" style="word-wrap&#58;break-word;margin&#58;0px;border&#58;0px;">
            //                <div class="ms-rtestate-read ms-rte-wpbox">
            //                  <div class="ms-rtestate-read b55b18a3-8a3b-453f-a714-7e8d803f4d30" id="div_b55b18a3-8a3b-453f-a714-7e8d803f4d30"></div>
            //                  <div class="ms-rtestate-read" id="vid_b55b18a3-8a3b-453f-a714-7e8d803f4d30" style="display&#58;none;"></div>
            //                </div>
            //              </div>
            //            </div>
            //          </td>
            //          <td class="ms-wiki-columnSpacing" style="width&#58;49.95%;">
            //            <div class="ms-rte-layoutszone-outer" style="width&#58;100%;">
            //              <div class="ms-rte-layoutszone-inner" style="word-wrap&#58;break-word;margin&#58;0px;border&#58;0px;">
            //                <div class="ms-rtestate-read ms-rte-wpbox">
            //                  <div class="ms-rtestate-read 0b2f12a4-3ab5-4a59-b2eb-275bbc617f95" id="div_0b2f12a4-3ab5-4a59-b2eb-275bbc617f95"></div>
            //                  <div class="ms-rtestate-read" id="vid_0b2f12a4-3ab5-4a59-b2eb-275bbc617f95" style="display&#58;none;"></div>
            //                </div>
            //              </div>
            //            </div>
            //          </td>
            //        </tr>
            //      </tbody>
            //    </table>
            //    <span id="layoutsData" style="display&#58;none;">true,false,2</span>
            //  </div>
            //</div>

            // Close all BR tags
            Regex brRegex = new Regex("<br>", RegexOptions.IgnoreCase);

            wikiField = brRegex.Replace(wikiField, "<br/>");

            XmlDocument xd = new XmlDocument();
            xd.PreserveWhitespace = true;
            xd.LoadXml(wikiField);

            // Sometimes the wikifield content seems to be surrounded by an additional div? 
            XmlElement layoutsTable = xd.SelectSingleNode("div/div/table") as XmlElement;
            if (layoutsTable == null)
            {
                layoutsTable = xd.SelectSingleNode("div/table") as XmlElement;
            }

            XmlElement layoutsZoneInner = layoutsTable.SelectSingleNode(string.Format("tbody/tr[{0}]/td[{1}]/div/div", row, col)) as XmlElement;
            // - space element
            XmlElement space = xd.CreateElement("p");
            XmlText text = xd.CreateTextNode(" ");
            space.AppendChild(text);

            // - wpBoxDiv
            XmlElement wpBoxDiv = xd.CreateElement("div");
            layoutsZoneInner.AppendChild(wpBoxDiv);

            if (addSpace)
            {
                layoutsZoneInner.AppendChild(space);
            }

            XmlAttribute attribute = xd.CreateAttribute("class");
            wpBoxDiv.Attributes.Append(attribute);
            attribute.Value = "ms-rtestate-read ms-rte-wpbox";
            attribute = xd.CreateAttribute("contentEditable");
            wpBoxDiv.Attributes.Append(attribute);
            attribute.Value = "false";
            // - div1
            XmlElement div1 = xd.CreateElement("div");
            wpBoxDiv.AppendChild(div1);
            div1.IsEmpty = false;
            attribute = xd.CreateAttribute("class");
            div1.Attributes.Append(attribute);
            attribute.Value = "ms-rtestate-read " + wpdNew.Id.ToString("D");
            attribute = xd.CreateAttribute("id");
            div1.Attributes.Append(attribute);
            attribute.Value = "div_" + wpdNew.Id.ToString("D");
            // - div2
            XmlElement div2 = xd.CreateElement("div");
            wpBoxDiv.AppendChild(div2);
            div2.IsEmpty = false;
            attribute = xd.CreateAttribute("style");
            div2.Attributes.Append(attribute);
            attribute.Value = "display:none";
            attribute = xd.CreateAttribute("id");
            div2.Attributes.Append(attribute);
            attribute.Value = "vid_" + wpdNew.Id.ToString("D");

            SP.ListItem listItem = webPartPage.ListItemAllFields;
            listItem["WikiField"] = xd.OuterXml;
            listItem.Update();
            web.Context.ExecuteQuery();

        }


        public string CreatePublishingPage(string siteUrl, string pagesListTitle, string layoutFileName, string pageTitle, string pageFileName, bool deleteIfExists)
        {
            string pageUrl = string.Empty;
            pageFileName = pageFileName.Replace(".aspx", "") + ".aspx";
            _logger.LogVerbose("Creation of publishing page: " + pageFileName);

            using (var context = GetContext(siteUrl))
            {
                SP.Web web = context.Web;
                SP.Web rootWeb = context.Site.RootWeb;
                context.Load(web, w => w.ServerRelativeUrl, w=>w.AllProperties);
                context.Load(rootWeb, w => w.ServerRelativeUrl);
                context.ExecuteQuery();


                SP.File existingFile = GetFile(context, pagesListTitle, null, pageFileName);

                if (deleteIfExists && existingFile!=null)
                {
                    {
                        _logger.LogVerbose("Deleting existing page...");
                        existingFile.DeleteObject();
                        context.ExecuteQuery();
                        existingFile = null;
                    }
                }


                if (existingFile == null)
                {
                    // Get Page Layout
                    string layoutPath = string.Format("{0}/_catalogs/masterpage/{1}.aspx", rootWeb.ServerRelativeUrl.TrimEnd('/'), layoutFileName.Replace(".aspx", ""));

                    _logger.LogVerbose("Looking for page layout: " + layoutPath);

                    SP.File pageFromDocLayout = rootWeb.GetFileByServerRelativeUrl(layoutPath);
                    SP.ListItem pageLayoutItem = pageFromDocLayout.ListItemAllFields;
                    context.Load(pageLayoutItem);
                    context.ExecuteQuery();

                    // Create Publishing Page
                    _logger.LogVerbose("Creating page...");
                    PublishingWeb publishingWeb = PublishingWeb.GetPublishingWeb(context, web);
                    PublishingPage page = publishingWeb.AddPublishingPage(new PublishingPageInformation
                    {
                        Name = pageFileName,
                        PageLayoutListItem = pageLayoutItem
                    });
                    context.ExecuteQuery();

                    // Set Page Title and Publish Page
                    _logger.LogVerbose("Setting page title and checking in...");
                    SP.ListItem pageItem = page.ListItem;
                    pageItem["Title"] = (string.IsNullOrEmpty(pageTitle) ? pageFileName : pageTitle);
                    pageItem.Update();
                    pageItem.File.CheckIn(String.Empty, SP.CheckinType.MajorCheckIn);
                    
                    context.ExecuteQuery();
                }
                else
                {
                    _logger.LogVerbose("Page already exists.");
                }
            }

            return pageUrl;
        }




        public void LoadWebPart(string siteUrl, string pageServerRelativeUrl, string webPartXml, string zoneId, int zoneIndex)
        {
            

            using (var context = GetContext(siteUrl))
            {
                _logger.LogVerbose("getting page file at: " + pageServerRelativeUrl);
                SP.File file = context.Web.GetFileByServerRelativeUrl(pageServerRelativeUrl);

                if (file == null) throw new Exception(string.Format("Page was not found at '{0}'", pageServerRelativeUrl));

                this.LoadWebPart(context, file, webPartXml, zoneId, zoneIndex);
            }
        }
        /// <summary>
        /// load web part into an existing page
        /// </summary>
        private void LoadWebPart(SP.ClientContext context, SP.File file, string webPartXml, string zoneId, int zoneIndex)
        {
            _logger.LogVerbose("Adding web part to page... ");

            try
            {
                file.CheckOut();

                LimitedWebPartManager wpmgr = file.GetLimitedWebPartManager(PersonalizationScope.Shared);
                WebPartDefinition wpd = wpmgr.ImportWebPart(webPartXml);
                wpmgr.AddWebPart(wpd.WebPart, zoneId, zoneIndex);

                file.CheckIn(String.Empty, SP.CheckinType.MajorCheckIn);
                context.ExecuteQuery();
            }
            catch
            {
                file.UndoCheckOut();
                context.ExecuteQuery();
                throw;
            }
            

        }

        public void AddWebPartIntoWikiPage(SP.ClientContext context, string pageUrl, string webPartXml)
        {
            var page = context.Web.GetFileByServerRelativeUrl(pageUrl);
            var webPartManager = page.GetLimitedWebPartManager(PersonalizationScope.Shared);

            var importedWebPart = webPartManager.ImportWebPart(webPartXml);
            var webPart = webPartManager.AddWebPart(importedWebPart.WebPart, "wpz", 0);
            context.Load(webPart);
            context.ExecuteQuery();

            string marker = String.Format("<div class=\"ms-rtestate-read ms-rte-wpbox\" contentEditable=\"false\"><div class=\"ms-rtestate-read {0}\" id=\"div_{0}\"></div><div style='display:none' id=\"vid_{0}\"></div></div>", webPart.Id);
            SP.ListItem item = page.ListItemAllFields;
            context.Load(item);
            context.ExecuteQuery();
            item["PublishingPageContent"] = marker;
            item.Update();
            context.ExecuteQuery();
        }

        
        
    }
}
