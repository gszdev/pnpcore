﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using PnP.Core.Test.Utilities;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace PnP.Core.Test.Base
{
    /// <summary>
    /// Test cases for the BatchClient class
    /// </summary>
    [TestClass]
    public class CSOMTests
    {
        private static readonly string WebTitleCsom = "<Request AddExpandoFieldTypeSuffix=\"true\" SchemaVersion=\"15.0.0.0\" LibraryVersion=\"16.0.0.0\" ApplicationName=\".NET Library\" xmlns=\"http://schemas.microsoft.com/sharepoint/clientquery/2009\"><Actions><ObjectPath Id=\"2\" ObjectPathId=\"1\" /><ObjectPath Id=\"4\" ObjectPathId=\"3\" /><Query Id=\"5\" ObjectPathId=\"3\"><Query SelectAllProperties=\"false\"><Properties><Property Name=\"Title\" ScalarProperty=\"true\" /></Properties></Query></Query></Actions><ObjectPaths><StaticProperty Id=\"1\" TypeId=\"{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}\" Name=\"Current\" /><Property Id=\"3\" ParentId=\"1\" Name=\"Web\" /></ObjectPaths></Request>";
        private static readonly string WebDescriptionCsom = "<Request AddExpandoFieldTypeSuffix=\"true\" SchemaVersion=\"15.0.0.0\" LibraryVersion=\"16.0.0.0\" ApplicationName=\".NET Library\" xmlns=\"http://schemas.microsoft.com/sharepoint/clientquery/2009\"><Actions><ObjectPath Id=\"2\" ObjectPathId=\"1\" /><ObjectPath Id=\"4\" ObjectPathId=\"3\" /><Query Id=\"5\" ObjectPathId=\"3\"><Query SelectAllProperties=\"false\"><Properties><Property Name=\"Description\" ScalarProperty=\"true\" /></Properties></Query></Query></Actions><ObjectPaths><StaticProperty Id=\"1\" TypeId=\"{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}\" Name=\"Current\" /><Property Id=\"3\" ParentId=\"1\" Name=\"Web\" /></ObjectPaths></Request>";


        [ClassInitialize]
        public static void TestFixtureSetup(TestContext testContext)
        {
            // Configure mocking default for all tests in this class, unless override by a specific test
            //TestCommon.Instance.Mocking = false;
            //TestCommon.Instance.GenerateMockingDebugFiles = true;
        }

        [TestMethod]
        public async Task SimplePropertyRequest()
        {
            //TestCommon.Instance.Mocking = false;
            using (var context = await TestCommon.Instance.GetContextAsync(TestCommon.TestSite))
            {
                var web = context.Web;

                // Get the title value via non CSOM
                await web.EnsurePropertiesAsync(p => p.Title);

                var apiCall = new ApiCall(WebTitleCsom);

                var response = await (web as Web).RawRequestAsync(apiCall, HttpMethod.Post);

                Assert.IsTrue(response.CsomResponseJson.Count > 0);
                Assert.IsTrue(response.CsomResponseJson[5].GetProperty("Title").GetString() == web.Title);
            }
        }

        [TestMethod]
        public async Task MultipleSimplePropertyRequests()
        {
            //TestCommon.Instance.Mocking = false;
            using (var context = await TestCommon.Instance.GetContextAsync(TestCommon.TestSite))
            {
                var web = context.Web;

                // Get the title value via non CSOM
                await web.EnsurePropertiesAsync(p => p.Title, p => p.Description);

                var apiCall1 = new ApiCall(WebTitleCsom);
                var apiCall2 = new ApiCall(WebDescriptionCsom);

                var batch = context.BatchClient.EnsureBatch();
                await (web as Web).RawRequestBatchAsync(batch, apiCall1, HttpMethod.Post);
                await (web as Web).RawRequestBatchAsync(batch, apiCall2, HttpMethod.Post);
                await context.ExecuteAsync(batch);

                var response1 = batch.Requests.First().Value;
                var response2 = batch.Requests.Last().Value;

                Assert.IsTrue(response1.CsomResponseJson.Count > 0);
                Assert.IsTrue(response2.CsomResponseJson.Count > 0);
                Assert.IsTrue(response1.CsomResponseJson[5].GetProperty("Title").GetString() == web.Title);
                Assert.IsTrue(response2.CsomResponseJson[5].GetProperty("Description").GetString() == web.Description);
            }
        }

        [TestMethod]
        public async Task CSOMPartitySubSequentModelLoads()
        {
            //TestCommon.Instance.Mocking = false;
            using (var context = await TestCommon.Instance.GetContextAsync(TestCommon.TestSite))
            {
                // Subsequent loads of the same model will "enrich" the loaded model
                var web = await context.Web.GetAsync(p => p.Title);
                Assert.IsTrue(web.IsPropertyAvailable(p => p.Title));
                Assert.IsFalse(web.IsPropertyAvailable(p => p.MasterUrl));

                await web.GetAsync(p => p.MasterUrl, p => p.AlternateCssUrl);
                Assert.IsTrue(web.IsPropertyAvailable(p => p.Title));
                Assert.IsTrue(web.IsPropertyAvailable(p => p.MasterUrl));
                Assert.IsTrue(web.IsPropertyAvailable(p => p.AlternateCssUrl));
                Assert.IsFalse(web.IsPropertyAvailable(p => p.CustomMasterUrl));
            }
        }

        [TestMethod]
        public async Task CSOMPartitySubSequentModelLoads2()
        {
            //TestCommon.Instance.Mocking = false;
            using (var context = await TestCommon.Instance.GetContextAsync(TestCommon.TestSite))
            {
                // Subsequent loads of the same model will "enrich" the loaded model:
                // Loading another child model with default properties
                var web = await context.Web.GetAsync(p => p.Title);
                Assert.IsTrue(web.IsPropertyAvailable(p => p.Title));
                Assert.IsFalse(web.IsPropertyAvailable(p => p.MasterUrl));

                await web.GetAsync(p => p.AssociatedOwnerGroup, p => p.AlternateCssUrl);
                Assert.IsTrue(web.IsPropertyAvailable(p => p.Title));
                Assert.IsTrue(web.IsPropertyAvailable(p => p.AssociatedOwnerGroup));
                Assert.IsTrue(web.IsPropertyAvailable(p => p.AlternateCssUrl));
                Assert.IsFalse(web.IsPropertyAvailable(p => p.CustomMasterUrl));

                Assert.IsTrue(web.AssociatedOwnerGroup.Requested);
                Assert.IsTrue(web.AssociatedOwnerGroup.IsPropertyAvailable(p => p.Title));
            }
        }

        [TestMethod]
        public async Task CSOMPartitySubSequentModelLoads3()
        {
            //TestCommon.Instance.Mocking = false;
            using (var context = await TestCommon.Instance.GetContextAsync(TestCommon.TestSite))
            {
                // Subsequent loads of the same model will "enrich" the loaded model:
                // Loading another child model via subsequent loads ==> CSOM does
                // not support this capability for child model loads, but it makes 
                // sense to behave identical to regular subsequent model loads
                var web = await context.Web.GetAsync(p => p.Title);
                Assert.IsTrue(web.IsPropertyAvailable(p => p.Title));
                Assert.IsFalse(web.IsPropertyAvailable(p => p.MasterUrl));

                await web.GetAsync(p => p.AssociatedOwnerGroup.LoadProperties(p=>p.Title), p => p.AlternateCssUrl);
                Assert.IsTrue(web.IsPropertyAvailable(p => p.Title));
                Assert.IsTrue(web.IsPropertyAvailable(p => p.AssociatedOwnerGroup));
                Assert.IsTrue(web.IsPropertyAvailable(p => p.AlternateCssUrl));
                Assert.IsFalse(web.IsPropertyAvailable(p => p.CustomMasterUrl));
                Assert.IsTrue(web.AssociatedOwnerGroup.Requested);
                Assert.IsTrue(web.AssociatedOwnerGroup.IsPropertyAvailable(p => p.Title));
                Assert.IsFalse(web.AssociatedOwnerGroup.IsPropertyAvailable(p => p.Description));

                await web.GetAsync(p => p.AssociatedOwnerGroup.LoadProperties(p => p.Description, p=>p.LoginName), p => p.AlternateCssUrl);
                Assert.IsTrue(web.IsPropertyAvailable(p => p.Title));
                Assert.IsTrue(web.IsPropertyAvailable(p => p.AssociatedOwnerGroup));
                Assert.IsTrue(web.IsPropertyAvailable(p => p.AlternateCssUrl));
                Assert.IsFalse(web.IsPropertyAvailable(p => p.CustomMasterUrl));
                Assert.IsTrue(web.AssociatedOwnerGroup.Requested);
                Assert.IsTrue(web.AssociatedOwnerGroup.IsPropertyAvailable(p => p.Title));
                Assert.IsTrue(web.AssociatedOwnerGroup.IsPropertyAvailable(p => p.Description));
                Assert.IsTrue(web.AssociatedOwnerGroup.IsPropertyAvailable(p => p.LoginName));
                Assert.IsFalse(web.AssociatedOwnerGroup.IsPropertyAvailable(p => p.AllowMembersEditMembership));

            }
        }

        [TestMethod]
        public async Task CSOMPartitySubSequentModelLoads4()
        {
            //TestCommon.Instance.Mocking = false;
            using (var context = await TestCommon.Instance.GetContextAsync(TestCommon.TestSite))
            {
                // Subsequent loads of the same model will "enrich" the loaded model:
                // Loading another model collection with default properties
                var web = await context.Web.GetAsync(p => p.Title);
                Assert.IsTrue(web.IsPropertyAvailable(p => p.Title));
                Assert.IsFalse(web.IsPropertyAvailable(p => p.MasterUrl));

                await web.GetAsync(p => p.Lists, p => p.AlternateCssUrl);
                Assert.IsTrue(web.IsPropertyAvailable(p => p.Title));
                Assert.IsTrue(web.IsPropertyAvailable(p => p.Lists));
                Assert.IsTrue(web.IsPropertyAvailable(p => p.AlternateCssUrl));
                Assert.IsFalse(web.IsPropertyAvailable(p => p.CustomMasterUrl));

                Assert.IsTrue(web.Lists.Requested);
                Assert.IsTrue(web.Lists.First().IsPropertyAvailable(p => p.Title));
            }
        }

        [TestMethod]
        public async Task CSOMPartitySubSequentModelLoads5()
        {
            //TestCommon.Instance.Mocking = false;
            using (var context = await TestCommon.Instance.GetContextAsync(TestCommon.TestSite))
            {
                // Subsequent loads of the same model will "enrich" the loaded model:
                // Loading another model collection with specific properties
                var web = await context.Web.GetAsync(p => p.Title);
                Assert.IsTrue(web.IsPropertyAvailable(p => p.Title));
                Assert.IsFalse(web.IsPropertyAvailable(p => p.MasterUrl));

                await web.GetAsync(p => p.Lists.LoadProperties(p=>p.Title), p => p.AlternateCssUrl);
                Assert.IsTrue(web.IsPropertyAvailable(p => p.Title));
                Assert.IsTrue(web.IsPropertyAvailable(p => p.Lists));
                Assert.IsTrue(web.IsPropertyAvailable(p => p.AlternateCssUrl));
                Assert.IsFalse(web.IsPropertyAvailable(p => p.CustomMasterUrl));

                Assert.IsTrue(web.Lists.Requested);
                Assert.IsTrue(web.Lists.First().IsPropertyAvailable(p => p.Title));
                Assert.IsFalse(web.Lists.First().IsPropertyAvailable(p => p.Description));
            }
        }

        [TestMethod]
        public async Task CSOMPartitySubSequentModelLoads6()
        {
            //TestCommon.Instance.Mocking = false;
            using (var context = await TestCommon.Instance.GetContextAsync(TestCommon.TestSite))
            {
                // Subsequent loads of the same model will "enrich" the loaded model:
                // Loading another model collection with specific properties followed by
                // a subsequent load requesting other properties ==> in this case the initially loaded collection model 
                // is replaced with the freshly loaded one, using the properties coming with the new one
                var web = await context.Web.GetAsync(p => p.Title);
                Assert.IsTrue(web.IsPropertyAvailable(p => p.Title));
                Assert.IsFalse(web.IsPropertyAvailable(p => p.MasterUrl));

                await web.GetAsync(p => p.Lists.LoadProperties(p => p.Title), p => p.AlternateCssUrl);
                Assert.IsTrue(web.IsPropertyAvailable(p => p.Title));
                Assert.IsTrue(web.IsPropertyAvailable(p => p.Lists));
                Assert.IsTrue(web.IsPropertyAvailable(p => p.AlternateCssUrl));
                Assert.IsFalse(web.IsPropertyAvailable(p => p.CustomMasterUrl));

                Assert.IsTrue(web.Lists.Requested);
                Assert.IsTrue(web.Lists.First().IsPropertyAvailable(p => p.Title));
                Assert.IsFalse(web.Lists.First().IsPropertyAvailable(p => p.Description));

                await web.GetAsync(p => p.Lists.LoadProperties(p => p.Description), p => p.MasterUrl);
                Assert.IsTrue(web.IsPropertyAvailable(p => p.Title));
                Assert.IsTrue(web.IsPropertyAvailable(p => p.Lists));
                Assert.IsTrue(web.IsPropertyAvailable(p => p.AlternateCssUrl));
                Assert.IsTrue(web.IsPropertyAvailable(p => p.MasterUrl));
                Assert.IsFalse(web.IsPropertyAvailable(p => p.CustomMasterUrl));

                Assert.IsTrue(web.Lists.Requested);
                Assert.IsTrue(web.Lists.First().IsPropertyAvailable(p => p.Description));
                Assert.IsFalse(web.Lists.First().IsPropertyAvailable(p => p.Title));
            }
        }

        [TestMethod]
        public async Task CSOMPartitySubSequentModelLoads7()
        {
            //TestCommon.Instance.Mocking = false;
            using (var context = await TestCommon.Instance.GetContextAsync(TestCommon.TestSite))
            {
                // This works different than CSOM, the requested list gets loaded into the model whereas with CSOM this is not the case
                var listA = await context.Web.Lists.GetByTitleAsync("Site Assets");
                var listB = await context.Web.Lists.GetByTitleAsync("Site Pages");

                Assert.IsTrue(listA.Requested);
                Assert.IsTrue(listB.Requested);
                Assert.IsTrue(context.Web.Lists.Count() == 2);

                var listC = await context.Web.Lists.GetByTitleAsync("Documents");
                Assert.IsTrue(listC.Requested);
                Assert.IsTrue(context.Web.Lists.Count() == 3);

                // Clear the collection so we can do a fresh load
                context.Web.Lists.Clear();
                Assert.IsFalse(context.Web.Lists.Requested);
                Assert.IsTrue(context.Web.Lists.Length == 0);

                // Populate the collection again
                await context.Web.Lists.GetAsync(p => p.TemplateType == ListTemplateType.DocumentLibrary);
                Assert.IsTrue(context.Web.Lists.Requested);
                Assert.IsTrue(context.Web.Lists.Length >= 3);
            }
        }
    }
}
