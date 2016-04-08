using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions
{
    internal static class SecurableObjectExtensions
    {
        public static void SetSecurity(this SecurableObject securable, TokenParser parser, ObjectSecurity security)
        {
            // If there's no role assignments we're returning
            if (security.RoleAssignments.Count == 0) return;

            var context = securable.Context as ClientContext;

            var groups = context.LoadQuery(context.Web.SiteGroups.Include(g => g.LoginName));
            var webRoleDefinitions = context.LoadQuery(context.Web.RoleDefinitions);

            context.ExecuteQueryRetry();

            securable.BreakRoleInheritance(security.CopyRoleAssignments, security.ClearSubscopes);

            foreach (var roleAssignment in security.RoleAssignments)
            {
                var roleAssignmentPrincipal = parser.ParseString(roleAssignment.Principal);
                Principal principal = groups.FirstOrDefault(g => g.LoginName == roleAssignmentPrincipal);
                if (principal == null)
                {
                    principal = context.Web.EnsureUser(roleAssignmentPrincipal);
                }

                if (principal != null)
                {
                var roleDefinitionBindingCollection = new RoleDefinitionBindingCollection(context);

                var roleDefinition = webRoleDefinitions.FirstOrDefault(r => r.Name == roleAssignment.RoleDefinition);

                if (roleDefinition != null)
                {
                    roleDefinitionBindingCollection.Add(roleDefinition);
                }
                securable.RoleAssignments.Add(principal, roleDefinitionBindingCollection);
            }
            }
            context.ExecuteQueryRetry();
        }

        public static ObjectSecurity GetSecurity(this SecurableObject securable, TokenParser parser)
        {
            ObjectSecurity security = null;
            using (var scope = new PnPMonitoredScope("Get Security"))
            {
                var context = securable.Context as ClientContext;

                context.Load(securable, sec => sec.HasUniqueRoleAssignments);
                var roleAssignments = context.LoadQuery(securable.RoleAssignments.Include(
                    r => r.Member.LoginName,
                    r => r.RoleDefinitionBindings.Include(
                        rdb => rdb.Name,
                        rdb => rdb.RoleTypeKind
                        )));
                context.ExecuteQueryRetry();

                if (securable.HasUniqueRoleAssignments)
                {
                    security = new ObjectSecurity();

                    foreach (var roleAssignment in roleAssignments)
                    {
                        if (roleAssignment.Member.LoginName != "Excel Services Viewers")
                        {
                            foreach (var roleDefinition in roleAssignment.RoleDefinitionBindings)
                            {
                                if (roleDefinition.RoleTypeKind != RoleType.Guest)
                                {
                                    security.RoleAssignments.Add(new Model.RoleAssignment()
                                    {
                                        Principal = parser.TokenizePrincipalLogin(roleAssignment.Member.LoginName),
                                        RoleDefinition = roleDefinition.Name
                                    });
                                }
                            }
                        }
                    }
                }
            }
            return security;
        }
    }
}
