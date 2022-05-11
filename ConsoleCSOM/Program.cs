using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
namespace ConsoleCSOM
{
    class SharepointInfo
    {
        public string RootSiteUrl { set; get; }
        public string SiteUrl { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
    }

    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                using (var clientContextHelper = new ClientContextHelper())
                {
                    //ClientContext ctx = GetContext(clientContextHelper);
                    //ctx.Load(ctx.Web);
                    //await ctx.ExecuteQueryAsync();
                    //Exercise 3 – Permission Inheritance
                    ////In the Finance and Accounting subsite, go to the List settings of the Accounts custom list and 
                    ////stop inheriting permissions.
                    //await StopInheritingPermissions(ctx, "Account");

                    // Add another user to the permission list with Design permissions.
                    //await GrantPermissions(ctx, "Account", "tu.nguyen.dev@devtusturu.onmicrosoft.com", "Design");

                    ////re - establish inheritance by selecting Delete unique permissions. 
                    //await ResetInheritingPermissions(ctx, "Account");
                    //Console.WriteLine($"Site {ctx.Web.Title}");

                    //Exercise 4 – Creating	Permission Levels and Groups
                    ClientContext ctxRoot = GetContextRootSite(clientContextHelper);
                    ctxRoot.Load(ctxRoot.Web);
                    //await CreateNewPermissionLevels(ctxRoot);
                    await CreateNewGroup(ctxRoot);
                    Console.WriteLine($"Site {ctxRoot.Web.Title}");
                }
                Console.WriteLine($"Press Any Key To Stop!");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        static ClientContext GetContext(ClientContextHelper clientContextHelper)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            var info = config.GetSection("SharepointInfo").Get<SharepointInfo>();
            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }

        static ClientContext GetContextRootSite(ClientContextHelper clientContextHelper)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            var info = config.GetSection("SharepointInfo").Get<SharepointInfo>();
            return clientContextHelper.GetContext(new Uri(info.RootSiteUrl), info.Username, info.Password);
        }

        private static async Task StopInheritingPermissions(ClientContext ctx, string listName)
        {
            List targetList = ctx.Web.Lists.GetByTitle(listName);
            ctx.Load(targetList);
            await ctx.ExecuteQueryAsync();

            //Stop Inheritance from parent
            targetList.BreakRoleInheritance(false, false);
            targetList.Update();
            await ctx.ExecuteQueryAsync();
        }

        private static async Task ResetInheritingPermissions(ClientContext ctx, string listName)
        {
            List targetList = ctx.Web.Lists.GetByTitle(listName);
            ctx.Load(targetList);
            await ctx.ExecuteQueryAsync();

            //Stop Inheritance from parent
            targetList.ResetRoleInheritance();
            targetList.Update();
            await ctx.ExecuteQueryAsync();
        }

        private static async Task GrantPermissions(ClientContext ctx, string listName, string nameOrEmail, string permission)
        {
            List targetList = ctx.Web.Lists.GetByTitle(listName);
            ctx.Load(targetList);
            await ctx.ExecuteQueryAsync();

            User user = ctx.Web.EnsureUser(nameOrEmail);
            ctx.Load(user);
            await ctx.ExecuteQueryAsync();

            RoleDefinition role = ctx.Web.RoleDefinitions.GetByName(permission);
            RoleDefinitionBindingCollection roleDb = new RoleDefinitionBindingCollection(ctx);
            roleDb.Add(role);

            targetList.RoleAssignments.Add(user, roleDb);
            targetList.Update();
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CreateNewPermissionLevels(ClientContext ctx)
        {
            RoleDefinitionCollection roleDefinitions = ctx.Web.RoleDefinitions;
            ctx.Load(roleDefinitions);
            await ctx.ExecuteQueryAsync();

            // Choose Permission and Set to Base Permission
            BasePermissions permissions = new BasePermissions();
            permissions.Set(PermissionKind.ManageLists);
            permissions.Set(PermissionKind.CreateAlerts);

            // Create New Custom Permission Level
            roleDefinitions.Add(new RoleDefinitionCreationInformation
            {
                Name = "Test Level",
                BasePermissions = permissions,
                Description = "Custom Permission Level With Manage Lists and Create Alerts"
            });
            ctx.Load(roleDefinitions);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CreateNewGroup(ClientContext ctx)
        {
            Group newgrp = ctx.Web.SiteGroups.Add(new GroupCreationInformation
            {
                Title = "Test Group",
                Description = "Test Group With Permission Is Test Level"
            });

            User user = ctx.Web.EnsureUser("tu.nguyen.dev@devtusturu.onmicrosoft.com");
            ctx.Load(user);
            await ctx.ExecuteQueryAsync();

            // Add User to Group
            newgrp.Users.Add(new UserCreationInformation
            {
                Email = user.Email,
                LoginName = user.LoginName,
                Title = user.Title,
            });
            ctx.ExecuteQuery();

            // Get the Role Definition (Permission Level)
            var targetPermissionLevel = ctx.Web.RoleDefinitions.GetByName("Test Level");
            ctx.Load(targetPermissionLevel);
            ctx.ExecuteQuery();

            // Add it to the Role Definition Binding Collection
            RoleDefinitionBindingCollection roleDb = new RoleDefinitionBindingCollection(ctx);
            roleDb.Add(ctx.Web.RoleDefinitions.GetByName("Test Level"));

            // Bind the Newly Created Permission Level to the new User Group
            ctx.Web.RoleAssignments.Add(newgrp, roleDb);

            ctx.Load(newgrp);
            await ctx.ExecuteQueryAsync();
        }
    }
}
