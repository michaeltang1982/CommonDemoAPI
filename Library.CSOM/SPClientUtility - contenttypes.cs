using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using Microsoft.SharePoint.Client.Publishing;


namespace Sierra.SharePoint.Library.CSOM
{
    public partial class SPClientUtility
    {


        /// <summary>
        /// create or update site column
        /// NOTE: DO NOT PUT THE VERSION ATTRIBUTE INTO FIELD XML
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="fieldName"></param>
        /// <param name="siteColumnXml"></param>
        /// <param name="addToDefaultView"></param>
        public void CreateUpdateSiteColumn(string siteUrl, string fieldName, string siteColumnXml, bool addToDefaultView)
        {
            using (var context = GetContext(siteUrl))
            {
                SP.Web web = context.Web;

                SP.Field field = this.GetFieldByInternalName(web, fieldName);

                if (field==null)    
                {
                    _logger.LogVerbose("creating new field");
                    field = web.Fields.AddFieldAsXml(siteColumnXml, addToDefaultView, SP.AddFieldOptions.AddFieldInternalNameHint);
                    field.UpdateAndPushChanges(true);
                }
                else
                {
                    _logger.LogVerbose("updating existing field");
                    field.SchemaXml = siteColumnXml;
                    field.UpdateAndPushChanges(true);
                }
                        

                //bool isTaxonomyField
                        //if (isTaxonomyField)
                            //{
                            //Guid termStoreId = Guid.Empty;
                            //Guid termSetId = Guid.Empty;
                            //GetTaxonomyFieldInfo(context, out termStoreId, out termSetId);

                            //// Retrieve as Taxonomy Field
                            //TaxonomyField taxonomyField = context.CastTo<TaxonomyField>(field);
                            //taxonomyField.SspId = termStoreId;
                            //taxonomyField.TermSetId = termSetId;
                            //taxonomyField.TargetTemplate = String.Empty;
                            //taxonomyField.AnchorId = Guid.Empty;
                            //taxonomyField.Update();
                        //}     
                  
                
                context.ExecuteQuery();
            }
        }




        public void CreateContentType(string siteUrl, string id, string contentTypeName, string group, string description, string documentTemplateUrl, List<string> fieldNames)
        {
            using (var context = GetContext(siteUrl))
            {
                SP.Web web = context.Web;

                SP.ContentType contentType = GetContentTypeById(context, web, id);

                if (contentType==null)
                {
                    _logger.LogVerbose("creating new content type...");
                    //create by Id
                    contentType = web.ContentTypes.Add(new SP.ContentTypeCreationInformation
                    {
                        Name = contentTypeName,
                        Id = id, 
                        Group = group,
                        Description = description
                    });
                    context.ExecuteQuery();

                    
                }

                _logger.LogVerbose("updating field references...");
                UpdateFieldRefs(context, web, contentType, fieldNames);

                if (!string.IsNullOrEmpty(documentTemplateUrl))
                {
                    _logger.LogVerbose("setting template url...");
                    contentType.DocumentTemplate = documentTemplateUrl;
                    contentType.Update(true);
                    context.ExecuteQuery();
                }
                
            }
        }






        private void UpdateFieldRefs(SP.ClientContext context, SP.Web rootWeb, SP.ContentType contentType, List<string> fieldNames)
        {
            SP.FieldLinkCollection fieldLinks = contentType.FieldLinks;
            context.Load(fieldLinks, f=>f.Include(l=>l.Name));
            context.ExecuteQuery();

            List<string> additionalFields = new List<string>();
            foreach (string fieldName in fieldNames) 
                if (fieldLinks.FirstOrDefault(f => f.Name == fieldName) == null) additionalFields.Add(fieldName);

            
            foreach (string fieldName in additionalFields)
            {
                SP.Field field = this.GetFieldByInternalName(rootWeb, fieldName);
                if (field == null) throw new Exception(string.Format("Field with name '{0}' was not found", fieldName));
                contentType.FieldLinks.Add(new SP.FieldLinkCreationInformation { Field = field });
                contentType.Update(true);
                context.ExecuteQuery();
            }

            
        }

        private void GetTaxonomyFieldInfo(SP.ClientContext clientContext, out Guid termStoreId, out Guid termSetId)
        {
            termStoreId = Guid.Empty;
            termSetId = Guid.Empty;

            //TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
            //TermStore termStore = session.GetDefaultSiteCollectionTermStore();
            //TermSetCollection termSets = termStore.GetTermSetsByName("SPSNL14", 1033);

            //clientContext.Load(termSets, tsc => tsc.Include(ts => ts.Id));
            //clientContext.Load(termStore, ts => ts.Id);
            //clientContext.ExecuteQuery();

            //termStoreId = termStore.Id;
            //termSetId = termSets.FirstOrDefault().Id;
        }



        internal SP.ContentType GetContentTypeById(SP.ClientContext context, SP.Web web, string id)
        {
            SP.ContentTypeCollection contentTypes = web.AvailableContentTypes;
            context.Load(contentTypes);
            context.ExecuteQuery();
            return contentTypes.FirstOrDefault(o => o.Id.StringValue == id);
        }

        internal SP.ContentType GetContentTypeByName(SP.ClientContext context, SP.Web web, string name)
        {
            SP.ContentTypeCollection contentTypes = web.AvailableContentTypes;
            context.Load(contentTypes);
            context.ExecuteQuery();
            return contentTypes.FirstOrDefault(o => o.Name == name);
        }

        internal SP.Field GetFieldByInternalName(SP.List list, string fieldName)
        {
            try
            {
                SP.Field field = list.Fields.GetByInternalNameOrTitle(fieldName);
                list.Context.Load(field);
                list.Context.ExecuteQuery();

                return field;
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("does not exist"))
                    //expected error where column does not exist
                    return null;
                else
                    throw;
            }
        }

        internal SP.Field GetFieldByInternalName(SP.Web web, string fieldName)
        {
            try
            {
                SP.ClientRuntimeContext context = web.Context;
                SP.Field field = web.Fields.GetByInternalNameOrTitle(fieldName);
                context.Load(web, w => w.Fields);
                context.Load(field, f => f.InternalName);
                context.ExecuteQuery();
                return field;
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("does not exist"))
                    //expected error where column does not exist
                    return null;
                else
                    throw;
            }

        }
    }
}
