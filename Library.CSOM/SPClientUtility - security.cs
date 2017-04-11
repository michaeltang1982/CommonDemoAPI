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
        



        public void CreateRoleDefinition(string siteUrl, string roleDefName, string roleDefDescription, List<string> permissions, bool deleteIfExists)
        {
            _logger.LogVerbose(string.Format("Create role definition '{0}' ", roleDefName));
            using (var context = GetContext(siteUrl))
            {
                SP.Web web = context.Site.RootWeb;
                context.Load(web);

                SP.RoleDefinitionCollection roleDefinitionCol = web.RoleDefinitions;
                context.Load(roleDefinitionCol);
                context.ExecuteQuery();

                var existingRoleDef = roleDefinitionCol.FirstOrDefault(p => p.Name == roleDefName);

                if (existingRoleDef != null && deleteIfExists)
                {
                    _logger.LogVerbose("Role definition already exists. Deleting...");
                    existingRoleDef.DeleteObject();
                    context.ExecuteQuery();

                    existingRoleDef = null;
                }

                if (existingRoleDef == null)
                {
                    _logger.LogVerbose("Role definition being created...");
                    var spRoleDef = new SP.RoleDefinitionCreationInformation();
                    var spBasePerm = new SP.BasePermissions();

                    foreach (string perm in permissions) { spBasePerm.Set((SP.PermissionKind)Enum.Parse(typeof(SP.PermissionKind), perm.Trim())); }

                    spRoleDef.Name = roleDefName;
                    spRoleDef.Description = roleDefDescription;
                    spRoleDef.BasePermissions = spBasePerm;
                    var roleDefinition = web.RoleDefinitions.Add(spRoleDef);
                    context.ExecuteQuery();
                }
            }

        }

        /// <summary>
        /// create a security group
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="groupName"></param>
        /// <param name="groupDescription"></param>
        public void CreateGroup(string siteUrl, string groupName, string groupDescription)
        {
            this.CreateGroup(siteUrl, groupName, groupDescription, false, false, false);
        }
        /// <summary>
        /// create a security group
        /// </summary>
        public void CreateGroup(string siteUrl, string groupName, string groupDescription, bool isAssociatedOwnerGroup, bool isAssociatedMemberGroup, bool isAssociatedVisitorGroup)
        {
            _logger.LogVerbose(string.Format("Create security group '{0}' at '{1}'", groupName, siteUrl));
            using (var context = GetContext(siteUrl))
            {
                SP.Group existingGroup = GetGroupByName(context, groupName);
                SP.Web web = context.Web;

                if(existingGroup==null)
                {
                    _logger.LogVerbose("Creating group...");
                    var newGroupInfo = new SP.GroupCreationInformation
                    {
                        Title = groupName,
                        Description = groupDescription

                    };


                   

                    existingGroup = web.SiteGroups.Add(newGroupInfo);

                                        
                    web.Update();

                    context.ExecuteQuery();
                }
                else
                {
                    _logger.LogVerbose("Group already exists");
                }

                //set association info -- ie thos groups which are associated wit owner, memeber and visitor default groups
                if (isAssociatedOwnerGroup) { web.AssociatedOwnerGroup = existingGroup; web.AssociatedOwnerGroup.Update(); }
                if (isAssociatedMemberGroup) { web.AssociatedMemberGroup = existingGroup; web.AssociatedMemberGroup.Update(); }
                if (isAssociatedVisitorGroup) { web.AssociatedVisitorGroup = existingGroup; web.AssociatedVisitorGroup.Update(); }

                if (isAssociatedOwnerGroup || isAssociatedMemberGroup || isAssociatedVisitorGroup)
                {
                    web.Update();
                    context.ExecuteQuery();
                }
            }
        }


        public void AssignGroupSecurityToSite(string siteUrl, string groupName, string roleName)
        {
            _logger.LogVerbose(string.Format("Assign to site '{0}': security group '{1}' with role '{2}'", siteUrl, groupName, roleName));
            using (var context = GetContext(siteUrl))
            {
                SP.Web web = context.Web;
                context.Load(web, w=>w.HasUniqueRoleAssignments);
                context.ExecuteQuery();

                this.AddGroupSecurity(web, web, groupName, roleName);

            }

        }

        public void RemoveGroupSecurityFromSite(string siteUrl, string groupName)
        {
            _logger.LogVerbose(string.Format("Remove from site '{0}': access for security group '{1}'", siteUrl, groupName));
            using (var context = GetContext(siteUrl))
            {

                SP.Web web = context.Web;
                
                context.Load(web, w=>w.HasUniqueRoleAssignments, w=>w.RoleAssignments.Include(r=>r.Member, r=>r.Member.PrincipalType));
                context.ExecuteQuery();
                
                if (!web.HasUniqueRoleAssignments) throw new Exception("Site does not have unique permissions");
                
                SP.RoleAssignmentCollection existingRoleAssignments = web.RoleAssignments;
                SP.RoleAssignment assignment = existingRoleAssignments.FirstOrDefault(r => r.Member.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.SharePointGroup && r.Member.Title == groupName);

                if (assignment != null)
                {
                    _logger.LogVerbose("Removing assignment for group ...");
                    web.RoleAssignments.GetByPrincipal(assignment.Member).DeleteObject();
                    web.Update();
                    context.ExecuteQuery();
                }
                else 
                { 
                    _logger.LogVerbose(string.Format("Group '{0}' is not assigned a role on site '{1}'", groupName, siteUrl)); 
                }
                
                


            }

        }



        public void AssignGroupSecurityToList(string siteUrl, string listTitle, string groupName, string roleName)
        {
            _logger.LogVerbose(string.Format("Assign to list '{0}': security group '{1}' with role '{2}'", listTitle, groupName, roleName));
            using (var context = GetContext(siteUrl))
            {
                SP.Web web = context.Web;

                SP.List list = this.GetListByTitle(context, listTitle, true);
                
                context.Load(list, l => l.HasUniqueRoleAssignments, l=>l.RoleAssignments);
                context.ExecuteQuery();

                this.AddGroupSecurity(web, list, groupName, roleName);
            }

        }

        public void RemoveGroupSecurityFromList(string siteUrl, string listTitle, string groupName)
        {
            _logger.LogVerbose(string.Format("Remove from list '{0}': access for security group '{1}'", listTitle, groupName));
            using (var context = GetContext(siteUrl))
            {

                SP.List list = this.GetListByTitle(context, listTitle, true);
                
                context.Load(list, l => l.HasUniqueRoleAssignments, l => l.RoleAssignments.Include(r => r.Member, r => r.Member.PrincipalType));
                context.ExecuteQuery();


                if (!list.HasUniqueRoleAssignments) throw new Exception("List does not have unique permissions");

                SP.RoleAssignmentCollection existingRoleAssignments = list.RoleAssignments;
                SP.RoleAssignment assignment = existingRoleAssignments.FirstOrDefault(r => r.Member.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.SharePointGroup && r.Member.Title == groupName);

                if (assignment != null)
                {
                    _logger.LogVerbose("Removing assignment for group ...");
                    list.RoleAssignments.GetByPrincipal(assignment.Member).DeleteObject();
                    list.Update();
                    context.ExecuteQuery();
                }
                else
                {
                    _logger.LogVerbose(string.Format("Group '{0}' is not assigned a role on list '{1}'", groupName, listTitle));
                }




            }

        }


        public void AssignGroupSecurityToListItem(string siteUrl, string listTitle, int itemId, string groupName, string roleName)
        {
            _logger.LogVerbose(string.Format("Assign to list item '{0}': security group '{1}' with role '{2}'", itemId, groupName, roleName));
            using (var context = GetContext(siteUrl))
            {
                SP.Web web = context.Web;

                SP.List list = this.GetListByTitle(context, listTitle, true);
                SP.ListItem listItem = list.GetItemById(itemId);

                context.Load(listItem, l => l.HasUniqueRoleAssignments, l => l.RoleAssignments);
                context.ExecuteQuery();

                this.AddGroupSecurity(web, listItem, groupName, roleName);
                
            }

        }


        public SP.Group GetGroupByName(SP.ClientContext context, string groupName)
        {
            SP.Group group = null;

            try
            {
                group = context.Web.SiteGroups.GetByName(groupName);
                context.ExecuteQuery();
            }
            catch (Exception ex)
            {
                if (!ex.Message.ToLower().Contains("cannot be found"))
                    throw;
                else
                    group = null;
            }

            return group;
        }



        private void AddGroupSecurity(SP.Web web, SecurableObject obj, string groupName, string roleName) 
        {
            if (!obj.HasUniqueRoleAssignments) throw new Exception("Object does not have unique permissions");

            SP.Group group = web.SiteGroups.GetByName(groupName);
            SP.RoleDefinition role = web.RoleDefinitions.GetByName(roleName);

            var roleBindings = new SP.RoleDefinitionBindingCollection(web.Context);

            roleBindings.Add(role);

            SP.RoleAssignmentCollection roleAssignmentCollection = obj.RoleAssignments;
            SP.RoleAssignment roleAssignment = roleAssignmentCollection.Add(group, roleBindings);

            web.Context.ExecuteQuery();
        }
    

        

    }
}
