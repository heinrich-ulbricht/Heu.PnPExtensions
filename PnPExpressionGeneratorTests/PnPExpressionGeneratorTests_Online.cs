using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core;
using PnPExtensions;
using System;

namespace PnPExpressionGeneratorTests
{
    [TestClass]
    public class PnPExpressionGeneratorTests_Online
    {
        private PnPExpressionGenerator gen;
        private ClientContext ctx;

        [TestInitialize]
        public void Initialize()
        {
            gen = new PnPExpressionGenerator();
            var authMgr = new AuthenticationManager();
            // set the environment variables HEUPNP_NAME and HEUPNP_PASSWORD and change the "siteUrl" variable to run the tests against your dev tenant
            var siteUrl = "https://heinrichulbricht.sharepoint.com/sites/dev";
            ctx = authMgr.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, Environment.GetEnvironmentVariable("HEUPNP_NAME"), Environment.GetEnvironmentVariable("HEUPNP_PASSWORD"));
        }

        [TestCleanup]
        public void Cleanup()
        {
            ctx.Dispose();
        }

        [TestMethod]
        public void TestPropertyOnClass()
        {
            var web = ctx.Web;
            var expr = gen.GetExpression(web, "Title");
            ctx.Load(web, expr);
            ctx.ExecuteQueryRetry();
            Assert.IsNotNull(web.Title);
        }

        [TestMethod]
        public void TestCollectionPropertyOnClass()
        {
            var web = ctx.Web;
            var expr = gen.GetExpression(web, "Lists.Title");
            ctx.Load(web, expr);
            ctx.ExecuteQueryRetry();
            Assert.IsNotNull(web.Lists[0].Title);
        }

        [TestMethod]
        public void TestPropertyOnCollection()
        {
            var lists = ctx.Web.Lists;
            var expr = gen.GetExpression(lists, "Title");
            ctx.Load(lists, expr);
            ctx.ExecuteQueryRetry();
            Assert.IsNotNull(lists[0].Title);
        }

        [TestMethod]
        public void TestSpecialPropertyOnCollection()
        {
            // assume Site Assets exists and there is an item...
            var items = ctx.Web.Lists.GetByTitle("Site Assets").GetItems(CamlQuery.CreateAllItemsQuery());
            var expr = gen.GetExpression(items, "FileRef");
            ctx.Load(items, expr);
            ctx.ExecuteQueryRetry();
            Assert.IsNotNull(items[0]["FileRef"]);
        }

        [TestMethod]
        public void TestCollectionCollectionPropertyOnCollection()
        {
            var web = ctx.Web;
            var expr = gen.GetExpression(web, "RoleAssignments.RoleDefinitionBindings.Name");
            ctx.Load(web, expr);
            ctx.ExecuteQueryRetry();
            Assert.IsNotNull(web.RoleAssignments[0].RoleDefinitionBindings[0].Name);
        }

    }
}
