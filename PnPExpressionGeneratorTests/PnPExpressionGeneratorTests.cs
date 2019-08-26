using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnPExtensions;

namespace PnPExpressionGeneratorTests
{
    [TestClass]
    public class PnPExpressionGeneratorTests
    {
        private PnPExpressionGenerator gen;
        
        [TestInitialize]
        public void Initialize()
        {
            gen = new PnPExpressionGenerator();
        }

        [TestMethod]
        public void TestPropertyOnClass()
        {
            var code = gen.GenerateExpressionCode(typeof(Web), "Title");
            Assert.AreEqual("a => a.Title", code);
        }

        [TestMethod]
        public void TestSpecialPropertyOnClass()
        {
            var code = gen.GenerateExpressionCode(typeof(Web), "PropertyKey");
            Assert.AreEqual("a => a[\"PropertyKey\"]", code);
        }

        [TestMethod]
        public void TestCollectionOnClass()
        {
            var code = gen.GenerateExpressionCode(typeof(Web), "Lists");
            Assert.AreEqual("a => Microsoft.SharePoint.Client.ClientObjectQueryableExtension.Include(a.Lists, b => b)", code);
        }

        [TestMethod]
        public void TestCollectionPropertyOnClass()
        {
            var code = gen.GenerateExpressionCode(typeof(Web), "Lists.Title");
            Assert.AreEqual("a => Microsoft.SharePoint.Client.ClientObjectQueryableExtension.Include(a.Lists, b => b.Title)", code);
        }

        [TestMethod]
        public void TestCollectionCollectionOnClass()
        {
            var code = gen.GenerateExpressionCode(typeof(Web), "Lists.RoleAssignments");
            Assert.AreEqual("a => Microsoft.SharePoint.Client.ClientObjectQueryableExtension.Include(a.Lists, b => Microsoft.SharePoint.Client.ClientObjectQueryableExtension.Include(b.RoleAssignments, c => c))", code);
        }

        [TestMethod]
        public void TestCollectionCollectionPropertyOnClass()
        {
            var code = gen.GenerateExpressionCode(typeof(Web), "Lists.RoleAssignments.Member");
            Assert.AreEqual("a => Microsoft.SharePoint.Client.ClientObjectQueryableExtension.Include(a.Lists, b => Microsoft.SharePoint.Client.ClientObjectQueryableExtension.Include(b.RoleAssignments, c => c.Member))", code);
        }

        [TestMethod]
        public void TestCollectionPropertyPropertyOnClass()
        {
            var code = gen.GenerateExpressionCode(typeof(Web), "Lists.DefaultView.Title");
            Assert.AreEqual("a => Microsoft.SharePoint.Client.ClientObjectQueryableExtension.Include(a.Lists, b => b.DefaultView.Title)", code);
        }        

        [TestMethod]
        public void TestPropertyOnCollection()
        {
            var code = gen.GenerateExpressionCode(typeof(ListItemCollection), "DisplayName");
            Assert.AreEqual("a => Microsoft.SharePoint.Client.ClientObjectQueryableExtension.Include(a, b => b.DisplayName)", code);
        }

        [TestMethod]
        public void TestSpecialPropertyOnCollection()
        {
            var code = gen.GenerateExpressionCode(typeof(ListItemCollection), "FileRef");
            Assert.AreEqual("a => Microsoft.SharePoint.Client.ClientObjectQueryableExtension.Include(a, b => b[\"FileRef\"])", code);
        }

        [TestMethod]
        public void TestCollectionCollectionPropertyOnCollection()
        {
            var code = gen.GenerateExpressionCode(typeof(Web), "RoleAssignments.RoleDefinitionBindings.Name");
            Assert.AreEqual("a => Microsoft.SharePoint.Client.ClientObjectQueryableExtension.Include(a.RoleAssignments, b => Microsoft.SharePoint.Client.ClientObjectQueryableExtension.Include(b.RoleDefinitionBindings, c => c.Name))", code);
        }
    }
}
