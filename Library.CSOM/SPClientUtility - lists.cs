using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client;
using System.IO.Packaging;
using System.IO;
using Sierra.NET.Core;

namespace Sierra.SharePoint.Library.CSOM
{
    /// <summary>
    /// list properties which can be updated
    /// </summary>
    public class ListUpdatableProperties
    {
        public bool BreakRoleInheritance { get; set; }
    }

    public partial class SPClientUtility
    {

        /// <summary>
        /// get titles of all the lists in the given site
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <returns></returns>
        public List<string> GetAllListTitles(string siteUrl)
        {

            using (var context = GetContext(siteUrl))
            {
                List<string> listTitles = new List<string>();
                SP.Web web = context.Web;
                ListCollection lists = context.Web.Lists;
                context.Load(lists, l => l.Include(t => t.Title));
                context.ExecuteQuery();

                foreach (var list in lists) listTitles.Add(list.Title);

                return listTitles;

            }


        }


        /// <summary>
        /// create a list - if it does not exist
        /// </summary>
        public void CreateList(string siteUrl, string templateName, string title, string description, bool deleteIfExists, List<string> fieldsXml)
        {
            _logger.LogVerbose("Creation of list with title: " + title);
            using (var context = GetContext(siteUrl))
            {
                _logger.LogVerbose("Check for existing list...");
                SP.List existingList = GetListByTitle(context, title);

                if (deleteIfExists && existingList != null)
                {
                    _logger.LogVerbose("Deleting existing list...");
                    existingList.DeleteObject();
                    context.ExecuteQuery();
                    existingList = null;
                }

                if (existingList == null)
                {
                    SP.Web web = context.Web;

                    _logger.LogVerbose("Creating new list...");
                    context.Load(web, s => s.ListTemplates);
                    context.ExecuteQuery();
                    var listCreationInfo = new SP.ListCreationInformation
                    {
                        Title = title,
                        Description = description
                    };
                    var listTemplate = web.ListTemplates.FirstOrDefault(template => template.Name == templateName);

                    if (listTemplate == null) throw new Exception(string.Format("List template {0} does not exist at {1}", templateName, siteUrl));

                    listCreationInfo.TemplateFeatureId = listTemplate.FeatureId;
                    listCreationInfo.TemplateType = listTemplate.ListTemplateTypeKind;

                    existingList = web.Lists.Add(listCreationInfo);
                    context.ExecuteQuery();

                }
                else
                {
                    _logger.LogVerbose(string.Format("List with title '{0}' already exists", title));
                }

                //check if fields need adding
                if (fieldsXml.Count() > 0) AddListFields(existingList, fieldsXml);
            }
        }

        /// <summary>
        /// update the properties of an existing list
        /// </summary>
        public void UpdateListProperties(string siteUrl, string listTitle, ListUpdatableProperties properties)
        {
            _logger.LogVerbose("Updating properties for list with title: " + listTitle);
            using (var context = GetContext(siteUrl))
            {
                SP.List list = GetListByTitle(context, listTitle, true);
                SP.Web web = context.Web;

                if (properties.BreakRoleInheritance)
                {
                    _logger.LogVerbose("breaking list role inheritance...");
                    list.BreakRoleInheritance(false, true);
                }

                context.ExecuteQuery();
            }


        }

        public void ApplyContentTypeToList(string siteUrl, string listTitle, string contentTypeName, bool isHidden)
        {
            _logger.LogVerbose(string.Format("Apply content type '{0}' to list '{1}'", contentTypeName, listTitle));
            using (var context = GetContext(siteUrl))
            {
                
                SP.List list = this.GetListByTitle(context, listTitle);
                if (list == null) throw new Exception(string.Format("List with title '{0}' does not exist", listTitle));

                SP.ContentType contentType = this.GetContentTypeByName(context, context.Web, contentTypeName);
                if (contentType == null) throw new Exception(string.Format("Content Type with name '{0}' does not exist", contentTypeName));

                SP.ContentTypeCollection listCTs = list.ContentTypes;
                context.Load(listCTs);
                context.ExecuteQuery();

                SP.ContentType contentTypeRef = listCTs.FirstOrDefault(c => c.Name == contentTypeName);

                if (contentTypeRef == null)
                {
                    list.ContentTypesEnabled = true;
                    contentTypeRef = list.ContentTypes.AddExistingContentType(contentType);
                                        
                }
                else 
                {
                    _logger.LogVerbose(string.Format("List '{0}' already contains content type '{1}'", listTitle, contentTypeName));
                }

                contentTypeRef.Hidden = isHidden;
                contentTypeRef.Update(false);
                list.Update();
                context.ExecuteQuery();
            }
 
        }

        public void RemoveContentTypesFromList(string siteUrl, string listTitle, List<string> contentTypeNames)
        {
            _logger.LogVerbose(string.Format("Remove content types '{0}' from list '{1}'", string.Join(",",contentTypeNames), listTitle));
            using (var context = GetContext(siteUrl))
            {
                SP.List list = this.GetListByTitle(context, listTitle);
                if (list == null) throw new Exception(string.Format("List with title '{0}' does not exist", listTitle));

                SP.ContentTypeCollection listCTs = list.ContentTypes;
                context.Load(listCTs);
                context.ExecuteQuery();

                var removeContentTypes = listCTs.Where(ct => contentTypeNames.Contains(ct.Name)).ToArray();

                foreach (SP.ContentType ct in removeContentTypes)
                {
                    ct.DeleteObject();
                }

                list.Update();
                context.ExecuteQuery();
            }
        }


        private void AddListFields(SP.List list, List<string> fieldsXml)
        {
            foreach (string fieldXml in fieldsXml)
            {
                string name = XmlHelper.ExtractAttributeFromXml(fieldXml, "name");
                _logger.LogVerbose("Processing list field: " + name);
            
                SP.Field listField = this.GetFieldByInternalName(list, name);
                SP.Field siteColumn = this.GetFieldByInternalName(list.ParentWeb, name);

                if (listField==null)
                {

                    if (siteColumn == null)
                    {
                        _logger.LogVerbose("New list field...");
                        list.Fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddFieldInternalNameHint);
                    }
                    else
                    {
                        _logger.LogVerbose("Add existing site column...");
                        list.Fields.Add(siteColumn);
                    }
                }
                else
                {
                    
                    if (siteColumn == null)
                    {
                        _logger.LogVerbose("Update existing list field...");
                        listField.SchemaXml = fieldXml;
                        listField.UpdateAndPushChanges(true);
                    }
                    else
                    {
                        _logger.LogVerbose("Update existing site column reference...");
                        //TODO
                    }
                }

                list.Context.ExecuteQuery();

            }
        }

        
        
        public void CreateView(string siteUrl, string listTitle, string viewTitle, List<string> viewFields, bool defaultView, bool paged, int rowLimit, string viewType, string query)
        {
            _logger.LogVerbose(string.Format("Create view '{0}' for list: {1}",viewTitle, listTitle));

            using (var context = GetContext(siteUrl))
            {
                SP.List list = this.GetListByTitle(context, listTitle);
                if (list == null) throw new Exception(string.Format("List with title '{0}' does not exist", listTitle));

                SP.ViewCollection views = list.Views;
                context.Load(views);
                context.ExecuteQuery();

                SP.View view = views.FirstOrDefault(v => v.Title == viewTitle);

                if (view != null)
                {
                    _logger.LogVerbose("View already exists. Deleting...");
                    view.DeleteObject();
                    context.ExecuteQuery();
                }

                _logger.LogVerbose("Creating view...");
                view = list.Views.Add(new SP.ViewCreationInformation
                {
                    Title = viewTitle,
                    ViewTypeKind = (SP.ViewType)Enum.Parse(typeof(SP.ViewType), viewType),
                    ViewFields = viewFields.ToArray(),
                    SetAsDefaultView = defaultView,
                    RowLimit = (uint)rowLimit,
                    PersonalView = false,
                    Paged = paged,
                    Query = query
                });            
                
                context.ExecuteQuery();

            }

        }

        public SP.List GetListByTitle(string siteUrl, string title)
        {
            using (var context = GetContext(siteUrl))
            {
                return GetListByTitle(context, title);
            }

        }

        public SP.List GetListByTitle(SP.ClientContext context, string title)
        {
            return this.GetListByTitle(context, title, false);
        }

        public SP.List GetListByTitle(SP.ClientContext context, string title, bool mustExist)
        {
            SP.List list = null;
            SP.Web web = context.Web;

            try
            {
                list = web.Lists.GetByTitle(title);
                context.ExecuteQuery();
            }
            catch (Exception ex)
            {
                if (mustExist) throw;

                //we are allowed to ignore list does not exist message
                if (!ex.Message.ToLower().Contains("does not exist"))
                    throw;
                else                
                    list = null;
                
            }
            
            return list;
        }

        public void UploadFile(string siteUrl, string sourceFilePath, string destinationListPath, string destinationFolderPath, string destinationFileName)
        {
            using (var context = GetContext(siteUrl))
            {
                SP.Web web = context.Web;

                _logger.LogVerbose(string.Format("Uploading file '{0}' ...", destinationFileName));


                SP.Folder folder = context.Web.GetFolderByServerRelativeUrl(GetListFolderUrl(destinationListPath, destinationFolderPath));


                SP.FileCreationInformation newFile = new SP.FileCreationInformation();
                newFile.Content = System.IO.File.ReadAllBytes(sourceFilePath);
                newFile.Url = destinationFileName;
                newFile.Overwrite = true;
                SP.File uploadFile = folder.Files.Add(newFile);
                web.Context.Load(uploadFile);
                web.Context.ExecuteQuery();


                context.ExecuteQuery();
            }
        }


        /// <summary>
        /// Get file url at the given location and folder
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="listTitle"></param>
        /// <param name="folderUrl"></param>
        /// <param name="filename"></param>
        /// <returns>server relative url or NULL if file does not exist</returns>
        public string GetExistingFileUrl(string siteUrl, string listTitle, string folderUrl, string filename)
        {
            using (var context = GetContext(siteUrl))
            {
                SP.File file = this.GetFile(context, listTitle, folderUrl, filename);
                if (file != null)
                    return file.ServerRelativeUrl;
                else
                    return null;
            }
        }


        


        /// <summary>
        /// Get file at the given location and folder
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="listTitle"></param>
        /// <param name="folderUrl"></param>
        /// <param name="filename"></param>
        /// <returns>SP.File or NULL</returns>
        public SP.File GetFile(string siteUrl, string listTitle, string folderUrl, string filename)
        {
            using (var context = GetContext(siteUrl))
            {
                return this.GetFile(context, listTitle, folderUrl, filename);
            }
        }

        /// <summary>
        /// Get file at the given location and folder
        /// </summary>
        /// <param name="context"></param>
        /// <param name="listTitle"></param>
        /// <param name="folderUrl"></param>
        /// <param name="filename"></param>
        /// <returns>SP.File or NULL</returns>
        public SP.File GetFile(SP.ClientContext context, string listTitle, string folderUrl, string filename)
        {
            
            SP.Web web = context.Web;
            SP.List list = web.Lists.GetByTitle(listTitle);
            context.Load(list, l => l.RootFolder, l => l.RootFolder.ServerRelativeUrl);
            context.ExecuteQuery();

            if (folderUrl == null) folderUrl = string.Empty;
            string fileUrl = string.Format("/{0}/{1}/{2}", list.RootFolder.ServerRelativeUrl.Trim('/'), folderUrl.Trim('/'), filename).Replace("//", "/");
            
            return this.GetFile(context, fileUrl);
            
        }


        /// <summary>
        /// Get file at the given location and folder
        /// </summary>
        /// <param name="context"></param>
        /// <param name="serverRelativeFileUrl"></param>
        /// <returns>SP.File or NULL</returns>
        public SP.File GetFile(SP.ClientContext context, string serverRelativeFileUrl)
        {

            SP.File file = null;
            
            try
            {
                file = context.Web.GetFileByServerRelativeUrl(serverRelativeFileUrl);
                context.Load(file, f=>f.CheckOutType);
                context.ExecuteQuery();
            }
            catch (Exception ex)
            {
                if (!ex.Message.ToLower().Contains("file not found"))
                    throw;
                else
                {
                    _logger.LogVerbose("No file found at " + serverRelativeFileUrl);
                    file = null;
                    
                }
            }

            return file;

        }

        /// <summary>
        /// get file at given location. if it exists, check it out
        /// </summary>
        /// <param name="context"></param>
        /// <param name="targetFileUrl"></param>
        /// <returns>Checked out file, or NULL</returns>
        public SP.File GetAndCheckOutFile(SP.ClientContext context, string targetFileUrl)
        {
            SP.File existingFile = this.GetFile(context, targetFileUrl);

            if (existingFile != null)
            {
                _logger.LogVerbose("found existing file at: " + targetFileUrl);
                if (existingFile.CheckOutType != CheckOutType.None)
                {
                    _logger.LogVerbose("attempting undo existing check out...");

                    try
                    {
                        existingFile.UndoCheckOut();
                        context.ExecuteQuery();
                    }
                    catch(Exception ex)
                    {
                        if (ex.Message.Contains("cannot discard check out"))
                        {
                            _logger.LogVerbose("undo check out failed. attempting delete...");
                            existingFile.DeleteObject();
                            context.ExecuteQuery();
                            existingFile = null;

                        }
                        else throw;
                    }
                }


                if (existingFile != null)
                {
                    _logger.LogVerbose("checking out existing file...");
                    existingFile.CheckOut();
                    context.ExecuteQuery();
                }
            }


            return existingFile;
        }
    }


}
