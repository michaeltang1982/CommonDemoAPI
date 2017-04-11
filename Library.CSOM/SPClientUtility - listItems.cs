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
    public partial class SPClientUtility
    {
        /// <summary>
        /// get list item in the given list
        /// </summary>
        /// <returns>list item or NULL</returns>
        public SP.ListItem GetListItemByTitle(SP.ClientContext context, SP.List list, string itemTitle)
        {
            return this.GetListItemByTitle(context, list, itemTitle, null);
        }

        /// <summary>
        /// get list item in the given list and folder
        /// </summary>
        /// <returns>list item or NULL</returns>
        public SP.ListItem GetListItemByTitle(SP.ClientContext context, SP.List list, string itemTitle, SP.Folder searchFolder)
        {
            var query = new SP.CamlQuery();
            query.ViewXml = string.Format("<View><Query><Where><Eq><FieldRef Name=\"Title\"/><Value Type=\"Text\">{0}</Value></Eq></Where></Query></View>", itemTitle);

            //restrict to a specific folder?
            if (searchFolder != null) query.FolderServerRelativeUrl = searchFolder.ServerRelativeUrl;

            var items = list.GetItems(query);
            context.Load(items, i => i.Include(m => m["Title"], m=>m.Id));
            context.ExecuteQuery();

            if (items.Count() > 0)
                return items[0];
            else
                return null;
        }


        /// <summary>
        /// Get list item ID associated with the file at the given server relative url
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="fileServerRelativeUrl"></param>        
        /// <returns>list item id or -1 if file does not exist</returns>
        public int GetListItemIdByFileUrl(string siteUrl, string fileServerRelativeUrl)
        {
            using (var context = GetContext(siteUrl))
            {
                SP.File file = this.GetFile(context, fileServerRelativeUrl);
                if (file != null)
                {
                    SP.ListItem item = file.ListItemAllFields;
                    context.Load(item);
                    context.ExecuteQuery();
                    return item.Id;
                }
                else
                    return -1;
            }
        }


        /// <summary>
        /// create new list item or document
        /// </summary>
        /// <returns>id of the list item</returns>
        public int CreateListItem(string siteUrl, string listTitle, string itemTitle, string folderUrl, string sourceFilePath, string newFileName, List<ItemProperty> itemProperties)
        {
            return this.CreateListItem(siteUrl, listTitle, itemTitle, folderUrl, sourceFilePath, newFileName, itemProperties, false);        
        }

        /// <summary>
        /// create new list item or document
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="listTitle">destination list title</param>
        /// <param name="itemTitle">(optional) title of new item</param>
        /// <param name="folderUrl">folder structure to place item into</param>
        /// <param name="sourceFilePath">source path of the uploading document</param>
        /// <param name="newFileName">(optional) new name for file once it is uploaded to library</param>
        /// <param name="itemProperties">additional properties to set for the list item</param>
        /// <returns>id of the list item</returns>
        public int CreateListItem(string siteUrl, string listTitle, string itemTitle, string folderUrl, string sourceFilePath, string newFileName, List<ItemProperty> itemProperties, bool breakRoleInheritance)
        {
            _logger.LogVerbose(string.Format("Creation of item in list: {0}...", listTitle));

            using (var context = this.GetContext(siteUrl))
            {
                SP.ListItem listItem = null;

                //list must exist
                SP.List list = this.GetListByTitle(context, listTitle, true);
                
                //get some extra info
                context.Load(list, l=>l.Title, l => l.RootFolder.ServerRelativeUrl, l=>l.BaseType);
                context.ExecuteQuery();

                bool isDocLib = list.BaseType == SP.BaseType.DocumentLibrary;
                SP.Folder folder = null;
                if (folderUrl == null) folderUrl = string.Empty; 

                //first check if folder is there..
                if (!string.IsNullOrEmpty(folderUrl))
                {
                    _logger.LogVerbose("ensure folder structure exists...");

                    //note: creating folders is different depending on type of list

                    if (isDocLib)
                    {
                        folder = this.CreateFolder(context, list.RootFolder, folderUrl);
                    }
                    else
                    {
                        folder = this.EnsureAndGetTargetFolder(context, list, folderUrl);
                    }

                }

                //now locate the item by title - if it exists
                if (!string.IsNullOrEmpty(itemTitle)) 
                { 
                    listItem = this.GetListItemByTitle(context, list, itemTitle, folder);
                    itemProperties.Add(new ItemProperty("Title", itemTitle, ""));
                }
                

                if (listItem == null)
                {
                    

                    if (!isDocLib)
                    {
                        _logger.LogVerbose("creating new list item...");
                        //normal list item
                        SP.ListItemCreationInformation itemCreateInfo = new SP.ListItemCreationInformation();
                        if (folder!=null) itemCreateInfo.FolderUrl = folder.ServerRelativeUrl;
                        listItem = list.AddItem(itemCreateInfo);                        
                    }
                    else
                    {
                        _logger.LogVerbose("creating new document...");
                        //this is a doc lib
                        listItem = CreateListItemDocLib(context, list, folderUrl, sourceFilePath, newFileName);

                    }
                }
                else
                {
                    _logger.LogVerbose("item or document exists.");

                    
                }

                if (isDocLib)
                {
                    EnsureCheckOut(context, listItem);
                }

                //update list item values                
                if (itemProperties.Count > 0)
                {
                    this.UpdateListItemMetadata(context, listItem, itemProperties);

                }

                if (isDocLib)
                {
                    _logger.LogVerbose("check-in the document...");
                    listItem.File.CheckIn("", CheckinType.OverwriteCheckIn);                    
                }

                if (breakRoleInheritance)
                {
                    _logger.LogVerbose("breaking list item role inheritance...");
                    listItem.BreakRoleInheritance(false, true);                    
                }

                //finally get id of the item
                context.Load(listItem, i => i.Id);
                context.ExecuteQuery();

                return listItem.Id;
            }

        }

        




        /// <summary>
        /// create a folder in the given list of any arbitrary nesting
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="listTitle"></param>
        /// <param name="fullFolderUrl">eg: /folder1/folder2/folder3 </param>
        /// <returns>new folder created</returns>
        public SP.Folder CreateFolder(string siteUrl, string listTitle, string fullFolderUrl)
        {
            _logger.LogVerbose(string.Format("Creation of folder {0} in list: {1}...", fullFolderUrl, listTitle));

            if (string.IsNullOrEmpty(fullFolderUrl)) throw new ArgumentNullException("fullFolderUrl");

            using (var context = this.GetContext(siteUrl))
            {
                //list must exist
                SP.List list = this.GetListByTitle(context, listTitle, true);
                return CreateFolder(context, list.RootFolder, fullFolderUrl);
            }


        }


        /// <summary>
        /// create actual list item in a document library (not a folder)
        /// </summary>
        /// <param name="folderUrl">(optional) folder name within the list</param>
        /// <param name="sourceFilePath">source path of file on client eg C:\Users\abc\test.docx</param>
        private SP.ListItem CreateListItemDocLib(SP.ClientContext context, SP.List list, string folderUrl, string sourceFilePath, string newFileName)
        {
            _logger.LogVerbose(string.Format("Uploading file '{0}' to list: {1}...", sourceFilePath, list.Title));
            var listRelativeUrl = list.RootFolder.ServerRelativeUrl;
            if (string.IsNullOrEmpty(newFileName))
            {
                //deduce filename from the source path
                newFileName = Path.GetFileName(sourceFilePath);
            }

            
            //build file upload path
            var targetFileUrl = string.Format("/{0}/{1}/{2}", listRelativeUrl.Trim(' ', '/'), folderUrl.Trim(' ', '/'), newFileName).Replace("//", "/");

            //check to see if file exists. if so, check it out
            SP.File existingFile = GetAndCheckOutFile(context, targetFileUrl);
            
            //do the upload of binaries
            _logger.LogVerbose("uploading binaries...");
            using (var fs = new FileStream(sourceFilePath, FileMode.Open))
            {
                SP.File.SaveBinaryDirect(context, targetFileUrl, fs, true);
            }

            //get the list item associated with the uploaded doc
            var uploadedFile = this.GetFile(context, targetFileUrl);
            var listItem = uploadedFile.ListItemAllFields;
            
            context.ExecuteQuery();

            return listItem;
        }

        

        /// <summary>
        /// update item metadata with values given as strings
        /// </summary>
        /// <param name="context"></param>
        /// <param name="item"></param>
        /// <param name="itemProperties"></param>
        private void UpdateListItemMetadata(SP.ClientContext context, SP.ListItem item, List<ItemProperty> itemProperties)
        {
            _logger.LogVerbose("Updating list item metadata...");

            string propertyName = string.Empty;

            try
            {
                //do taxonomy fields first
                foreach (ItemProperty property in itemProperties.Where(i => i.PropertyType.Name.Contains("TaxonomyFieldValue")))
                {
                    if (!string.IsNullOrEmpty(property.PropertyValue))
                    {
                        propertyName = property.PropertyName;
                        this.SetManagedMetaDataField(context, item.ParentList, item, property.PropertyName, property.PropertyValue);
                    }
                }

                //now do the rest
                foreach (ItemProperty property in itemProperties.Where(i => !i.PropertyType.Name.Contains("TaxonomyFieldValue")))
                {
                    if (!string.IsNullOrEmpty(property.PropertyValue))
                    {
                        propertyName = property.PropertyName;

                        object val = property.PropertyValue;
                        //bool setFieldValue = true;

                        if (property.PropertyType == typeof(DateTime))
                        {
                            //date
                            val = DateTime.Parse(val.ToString());
                        }
                        else if (property.PropertyType == typeof(SP.FieldUserValue))
                        {
                            //person or group
                            SP.User _newUser = context.Web.EnsureUser(val.ToString());
                            context.Load(_newUser);
                            context.ExecuteQuery();

                            SP.FieldUserValue _userValue = new SP.FieldUserValue();
                            _userValue.LookupId = _newUser.Id;
                            val = _userValue;

                        }

                        item[property.PropertyName] = val;

                    }
                }
                item.Update();
                context.ExecuteQuery();
            }
            catch 
            {
                _logger.LogWarning("Error setting property: " + propertyName);
                throw;
            }
        }

        /// <summary>
        /// break the role inheritance of the given list item
        /// </summary>
        public void BreakRoleInheritanceOfListItem(string siteUrl, string listTitle, int itemId)
        {
            _logger.LogVerbose("Breaking role inheritance of list item: " + itemId.ToString());
            using (var context = this.GetContext(siteUrl))
            {
                //list must exist
                SP.List list = this.GetListByTitle(context, listTitle, true);
                SP.ListItem item = list.GetItemById(itemId);
                item.BreakRoleInheritance(false, true);
                context.ExecuteQuery();
            }
        }

        /// <summary>
        /// recursively create folder structure starting from the list root
        /// </summary>
        /// <param name="context"></param>
        /// <param name="parentFolder"></param>
        /// <param name="fullFolderUrl"></param>
        /// <returns>new folder</returns>
        private Folder CreateFolder(SP.ClientContext context, SP.Folder parentFolder, string fullFolderUrl)
        {
            var folderUrls = fullFolderUrl.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string folderUrl = folderUrls[0];
            var curFolder = parentFolder.Folders.Add(folderUrl);
            context.Load(curFolder);
            context.ExecuteQuery();

            if (folderUrls.Length > 1)
            {
                var subFolderUrl = string.Join("/", folderUrls, 1, folderUrls.Length - 1);
                return CreateFolder(context, curFolder, subFolderUrl);
            }
            return curFolder;
        }

        /// <summary>
        /// make sure that if the file associated with the listitem needs checkout, it should be checked out
        /// </summary>
        /// <param name="context"></param>
        /// <param name="listItem"></param>
        private void EnsureCheckOut(SP.ClientContext context, SP.ListItem listItem)
        {
            _logger.LogVerbose("ensuring file is checked out...");
            SP.File file = listItem.File;
            context.Load(file, f => f.CheckOutType);
            context.ExecuteQuery();

            if (file.CheckOutType == CheckOutType.None)
            {
                listItem.File.CheckOut();
                context.ExecuteQuery();
            }

        }

        /// <summary>
        /// Will ensure nested folder creation if folders in folderPath don't exist.
        /// </summary>
        /// <param name="context">Loaded SharePoint Client Context</param>
        /// <param name="list">Document Library SharePoint List Object</param>
        /// <param name="folderPath">folder url such as /Folder1/Folder2/..</param>
        /// <returns>Last ChildFolder as target</returns>
        private Folder EnsureAndGetTargetFolder(ClientContext context, List list, string folderUrl)
        {
            var folderUrls = folderUrl.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            return this.EnsureAndGetTargetFolder(context, list, folderUrls);
        }

        /// <summary>
        /// Will ensure nested folder creation if folders in folderPath don't exist.
        /// </summary>
        /// <param name="context">Loaded SharePoint Client Context</param>
        /// <param name="list">Document Library SharePoint List Object</param>
        /// <param name="folderPath">List of strings ParentFolder, ChildFolder, ...</param>
        /// <returns>Last ChildFolder as target</returns>
        private Folder EnsureAndGetTargetFolder(ClientContext context, List list, string[] folderPath)
        {
            Folder returnFolder = list.RootFolder;
            if (folderPath != null && folderPath.Length > 0)
            {
                Web web = context.Web;
                Folder currentFolder = list.RootFolder;
                context.Load(web, t => t.Url);
                context.Load(currentFolder);
                context.ExecuteQuery();
                foreach (string folderPointer in folderPath)
                {
                    FolderCollection folders = currentFolder.Folders;
                    context.Load(folders);
                    context.ExecuteQuery();

                    bool folderFound = false;
                    foreach (Folder existingFolder in folders)
                    {
                        if (existingFolder.Name.Equals(folderPointer, StringComparison.InvariantCultureIgnoreCase))
                        {
                            folderFound = true;
                            currentFolder = existingFolder;
                            break;
                        }
                    }

                    if (!folderFound)
                    {
                        ListItemCreationInformation itemCreationInfo = new ListItemCreationInformation();
                        itemCreationInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                        itemCreationInfo.LeafName = folderPointer.Trim();
                        itemCreationInfo.FolderUrl = currentFolder.ServerRelativeUrl;
                        ListItem folderItemCreated = list.AddItem(itemCreationInfo);
                        folderItemCreated["Title"] = folderPointer;
                        folderItemCreated.Update();
                        context.Load(folderItemCreated, f => f.Folder);
                        context.ExecuteQuery();
                        currentFolder = folderItemCreated.Folder;

                    }
                }
                returnFolder = currentFolder;
            }
            return returnFolder;
        }

    }


}
