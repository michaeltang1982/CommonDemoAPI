using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using Microsoft.SharePoint.Client.Publishing;
using Microsoft.SharePoint.Client.Taxonomy;


namespace Sierra.SharePoint.Library.CSOM
{
    public partial class SPClientUtility
    {

        //protected void OnExecute(paramsobject[] args)
        //{
        //    try
        //    {

        //        string webUrl = GetParameter<string>(args, 1);

        //        string xmlData = GetParameter<string>(args, 2);

        //        TermStoreInfo termStoreInfo = TermStoreInfo.FromXml(xmlData);

        //        using (var context = new SharePointWebContext(webUrl))
        //        {

        //            Microsoft.SharePoint.Client.Web web = context.Web;
        //            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(web.Context);
        //            TermStore termStore = taxonomySession.TermStores.GetByName(termStoreInfo.Name);
        //            web.Context.Load(termStore);
        //            web.Context.ExecuteQuery();

        //            if (termStore == null)
        //            {
        //                Log.ErrorFormat("The Taxonomy Service is offline or unavailable.");
        //                return;
        //            }
        //            CreateGroups(termStoreInfo, termStore, web);
        //        }

        //    }
        //    catch (Exception x)
        //    {
        //        Log.Exception(x);
        //    }
        //}



        //protected void CreateGroups(TermStoreInfo termstoreInfo, TermStore termStore, Microsoft.SharePoint.Client.Web web)
        //{
        //    if (termstoreInfo != null && termstoreInfo.Groups.Count > 0)
        //    {
        //        foreach (GroupInfo groupInfo in termstoreInfo.Groups)
        //        {
        //            TermGroup termGroup = default(TermGroup);

        //            try
        //            {
        //                termGroup = termStore.Groups.GetByName(groupInfo.Name);
        //                web.Context.Load(termGroup);
        //                web.Context.ExecuteQuery();
        //            }
        //            catch (Exception)
        //            {
        //                termGroup = null;
        //            }

        //            if (termGroup == null)
        //            {
        //                termGroup = termStore.CreateGroup(groupInfo.Name, Guid.NewGuid());
        //                termStore.CommitAll();
        //                web.Context.ExecuteQuery();
        //            }

        //            CreateTermsets(groupInfo, termGroup, termStore, web);

        //        }
        //    }
        //}

        //protected void CreateTermsets(GroupInfo groupInfo, TermGroup termGroup, TermStore termStore, Microsoft.SharePoint.Client.Web web)
        //{

        //    foreach (TermSetInfo termsetInfo in groupInfo.TermSets)
        //    {
        //        TermSet termSet = default(TermSet);

        //        try
        //        {
        //            termSet = termGroup.TermSets.GetByName(termsetInfo.Name);
        //            web.Context.Load(termSet);
        //            web.Context.ExecuteQuery();
        //        }
        //        catch (Exception)
        //        {
        //            termSet = null;
        //        }

        //        if (termSet == null)
        //        {

        //            termSet = termGroup.CreateTermSet(termsetInfo.Name, Guid.NewGuid(), 1033);
        //            termStore.CommitAll();
        //            web.Context.ExecuteQuery();
        //        }

        //        CreateTerms(termsetInfo, termSet, termStore, web);
        //    }

        //}

        //protected void CreateTerms(TermSetInfo termSetInfo, TermSet termSet, TermStore termStore, Microsoft.SharePoint.Client.Web web)
        //{

        //    foreach (TermInfo termInfo in termSetInfo.Terms)
        //    {

        //        Term term = default(Term);

        //        try
        //        {
        //            term = termSet.Terms.GetByName(termInfo.Name);
        //            web.Context.Load(term);
        //            web.Context.ExecuteQuery();
        //        }
        //        catch (Exception)
        //        {
        //            term = null;
        //        }

        //        if (term == null)
        //        {
        //            term = termSet.CreateTerm(termInfo.Name, 1033, Guid.NewGuid());
        //            termStore.CommitAll();
        //            web.Context.ExecuteQuery();
        //        }

        //        CreateTerms(termInfo, term, termStore, web);
        //    }
        //}

        //protected void CreateTerms(TermInfo termInfo, Term term, TermStore termStore, Microsoft.SharePoint.Client.Web web)
        //{
        //    foreach (TermInfo childTermInfo in termInfo.ChildTerms)
        //    {
        //        Term childTerm = default(Term);

        //        try
        //        {
        //            childTerm = term.Terms.GetByName(childTermInfo.Name);
        //            web.Context.Load(childTerm);
        //            web.Context.ExecuteQuery();
        //        }
        //        catch (Exception)
        //        {
        //            childTerm = null;
        //        }

        //        if (childTerm == null)
        //        {
        //            childTerm = term.CreateTerm(childTermInfo.Name, 1033, Guid.NewGuid());
        //            termStore.CommitAll();
        //            web.Context.ExecuteQuery();
        //        }

        //        CreateTerms(childTermInfo, childTerm, termStore, web);
        //    }
        //}


        

        public string GetTermIdForTerm(ClientContext context, string term, Guid termSetId)
        {
            string termId = string.Empty;

            TaxonomySession tSession = TaxonomySession.GetTaxonomySession(context);
            TermStore ts = tSession.GetDefaultSiteCollectionTermStore();
            TermSet tset = ts.GetTermSet(termSetId);

            LabelMatchInformation lmi = new LabelMatchInformation(context);

            lmi.Lcid = 1033;
            lmi.TrimUnavailable = true;
            lmi.TermLabel = term;

            TermCollection termMatches = tset.GetTerms(lmi);
            context.Load(tSession);
            context.Load(ts);
            context.Load(tset);
            context.Load(termMatches);

            context.ExecuteQuery();

            if (termMatches != null && termMatches.Count() > 0)
                termId = termMatches.First().Id.ToString();

            return termId;

        }

        private void SetManagedMetaDataField(ClientContext context, SP.List list, SP.ListItem item, string fieldName, string term)
        {
            try
            {
                FieldCollection fields = list.Fields;
                Field field = fields.GetByInternalNameOrTitle(fieldName);

                context.Load(fields);
                context.Load(field);
                context.ExecuteQuery();

                TaxonomyField txField = context.CastTo<TaxonomyField>(field);
                string termId = GetTermIdForTerm(context, term, txField.TermSetId);

                if (string.IsNullOrEmpty(termId)) throw new Exception("Term value not recognised");
                
                TaxonomyFieldValueCollection termValues = null;
                TaxonomyFieldValue termValue = null;

                string termValueString = string.Empty;

                if (txField.AllowMultipleValues)
                {

                    termValues = item[fieldName] as TaxonomyFieldValueCollection;
                    foreach (TaxonomyFieldValue tv in termValues)
                    {
                        termValueString += tv.WssId + ";#" + tv.Label + "|" + tv.TermGuid + ";#";
                    }

                    termValueString += "-1;#" + term + "|" + termId;
                    termValues = new TaxonomyFieldValueCollection(context, termValueString, txField);
                    txField.SetFieldValueByValueCollection(item, termValues);

                }
                else
                {
                    termValue = new TaxonomyFieldValue();
                    termValue.Label = term;
                    termValue.TermGuid = termId;
                    termValue.WssId = -1;
                    txField.SetFieldValueByValue(item, termValue);
                }

                item.Update();
                context.Load(item);
                context.ExecuteQuery();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, string.Format("Error setting taxonomy field value: {0} = '{1}'. ", fieldName, term));
                throw;
            }
        }


    }
}
