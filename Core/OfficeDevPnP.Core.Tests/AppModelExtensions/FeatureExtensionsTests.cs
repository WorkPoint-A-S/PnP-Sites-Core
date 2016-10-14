﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Tests;
using Microsoft.Online.SharePoint.TenantAdministration;
using OfficeDevPnP.Core;
using System.Configuration;
using OfficeDevPnP.Core.Entities;

namespace Microsoft.SharePoint.Client.Tests
{
    [TestClass()]
    public class FeatureExtensionsTests
    {
        const string TEST_CATEGORY = "Feature Extensions";
        private ClientContext clientContext;
        private Guid sp2007WorkflowSiteFeatureId = new Guid("c845ed8d-9ce5-448c-bd3e-ea71350ce45b");
        private Guid contentOrganizerWebFeatureId = new Guid("7ad5272a-2694-4349-953e-ea5ef290e97c");
        private Guid publishingSiteFeatureId = new Guid("f6924d36-2fa8-4f0b-b16d-06b7250180fa");
        private Guid publishingWebFeatureId = new Guid("94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb");

        private static string sitecollectionNamePrefix = "TestPnPSC_123456789_";
        private string sitecollectionName = "";

        #region Test initialize and cleanup
        [TestInitialize()]
        public void Initialize()
        {
            clientContext = TestCommon.CreateClientContext();
            sitecollectionName = sitecollectionNamePrefix + Guid.NewGuid().ToString();
        }

        [TestCleanup()]
        public void Cleanup()
        {
            clientContext.Dispose();
        }

        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            using (var tenantContext = TestCommon.CreateTenantClientContext())
            {
                CleanupAllTestSiteCollections(tenantContext);
            }
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
            using (var tenantContext = TestCommon.CreateTenantClientContext())
            {
                CleanupAllTestSiteCollections(tenantContext);
            }
        }

#if !ONPREMISES
        internal static string CreateTestSiteCollection(Tenant tenant, string sitecollectionName)
        {
            try
            {
                string devSiteUrl = ConfigurationManager.AppSettings["SPODevSiteUrl"];
                string siteToCreateUrl = GetTestSiteCollectionName(devSiteUrl, sitecollectionName);

                string siteOwnerLogin = ConfigurationManager.AppSettings["SPOUserName"];
                if (TestCommon.AppOnlyTesting())
                {
                    using (var clientContext = TestCommon.CreateClientContext())
                    {
                        List<UserEntity> admins = clientContext.Web.GetAdministrators();
                        siteOwnerLogin = admins[0].LoginName.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries)[2];
                    }
                }

                SiteEntity siteToCreate = new SiteEntity()
                {
                    Url = siteToCreateUrl,
                    Template = "STS#0",
                    Title = "Test",
                    Description = "Test site collection",
                    SiteOwnerLogin = siteOwnerLogin,
                };

                tenant.CreateSiteCollection(siteToCreate, false, true);
                return siteToCreateUrl;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw;
            }
        }

        private static void CleanupAllTestSiteCollections(ClientContext tenantContext)
        {
            var tenant = new Tenant(tenantContext);

            var siteCols = tenant.GetSiteCollections();

            foreach (var siteCol in siteCols)
            {
                if (siteCol.Url.Contains(sitecollectionNamePrefix))
                {
                    try
                    {
                        // Drop the site collection from the recycle bin
                        if (tenant.CheckIfSiteExists(siteCol.Url, "Recycled"))
                        {
                            tenant.DeleteSiteCollectionFromRecycleBin(siteCol.Url, false);
                        }
                        else
                        {
                            // Eat the exceptions: would occur if the site collection is already in the recycle bin.
                            try
                            {
                                // ensure the site collection in unlocked state before deleting
                                tenant.SetSiteLockState(siteCol.Url, SiteLockState.Unlock);
                            }
                            catch { }

                            // delete the site collection, do not use the recyle bin
                            tenant.DeleteSiteCollection(siteCol.Url, false);
                        }
                    }
                    catch (Exception ex)
                    {
                        // eat all exceptions
                        Console.WriteLine(ex.ToString());
                    }
                }
            }
        }

#else
        private static string CreateTestSiteCollection(Tenant tenant, string sitecollectionName)
        {
            string devSiteUrl = ConfigurationManager.AppSettings["SPODevSiteUrl"];

            string siteOwnerLogin = string.Format("{0}\\{1}", ConfigurationManager.AppSettings["OnPremDomain"], ConfigurationManager.AppSettings["OnPremUserName"]);
            if (TestCommon.AppOnlyTesting())
            {
                using (var clientContext = TestCommon.CreateClientContext())
                {
                    List<UserEntity> admins = clientContext.Web.GetAdministrators();
                    siteOwnerLogin = admins[0].LoginName.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries)[1];
                }
            }

            string siteToCreateUrl = GetTestSiteCollectionName(devSiteUrl, sitecollectionName);
            SiteEntity siteToCreate = new SiteEntity()
            {
                Url = siteToCreateUrl,
                Template = "STS#0",
                Title = "Test",
                Description = "Test site collection",
                SiteOwnerLogin = siteOwnerLogin,
            };

            tenant.CreateSiteCollection(siteToCreate);
            return siteToCreateUrl;
        }

        private static void CleanupAllTestSiteCollections(ClientContext tenantContext)
        {
            string devSiteUrl = ConfigurationManager.AppSettings["SPODevSiteUrl"];

            var tenant = new Tenant(tenantContext);
            try
            {
                using (ClientContext cc = tenantContext.Clone(devSiteUrl))
                {
                    var sites = cc.Web.SiteSearch();

                    foreach (var site in sites)
                    {
                        if (site.Url.ToLower().Contains(sitecollectionNamePrefix.ToLower()))
                        {
                            tenant.DeleteSiteCollection(site.Url);
                        }
                    }
                }
            }
            catch
            { }
        }

        private void CleanupCreatedTestSiteCollections(ClientContext tenantContext)
        {
            string devSiteUrl = ConfigurationManager.AppSettings["SPODevSiteUrl"];
            String testSiteCollection = GetTestSiteCollectionName(devSiteUrl, sitecollectionName);

            //Ensure the test site collection was deleted and removed from recyclebin
            var tenant = new Tenant(tenantContext);
            try
            {
                tenant.DeleteSiteCollection(testSiteCollection);
            }
            catch
            { }
        }
#endif

        private static string GetTestSiteCollectionName(string devSiteUrl, string siteCollection)
        {
            Uri u = new Uri(devSiteUrl);
            string host = String.Format("{0}://{1}", u.Scheme, u.DnsSafeHost);

            string path = u.AbsolutePath;
            if (path.EndsWith("/"))
            {
                path = path.Substring(0, path.Length - 1);
            }
            path = path.Substring(0, path.LastIndexOf('/'));

            return string.Format("{0}{1}/{2}", host, path, siteCollection);
        }
        #endregion

        #region Feature activation tests
        [TestMethod()]
        [Timeout(45 * 60 * 1000)]
        public void PublishingFeatureActivationTest()
        {
            using (var tenantContext = TestCommon.CreateTenantClientContext())
            {
                var tenant = new Tenant(tenantContext);

                //Create site collection test
                Console.WriteLine("CreateDeleteSiteCollectionTest: step 1");
                string siteToCreateUrl = CreateTestSiteCollection(tenant, sitecollectionName);
                Console.WriteLine("CreateDeleteSiteCollectionTest: step 1.1");

                using (var clientContext = TestCommon.CreateClientContext(siteToCreateUrl))
                {
                    // Activate
                    clientContext.Site.ActivateFeature(publishingSiteFeatureId);
                    Assert.IsTrue(clientContext.Site.IsFeatureActive(publishingSiteFeatureId));
                    Console.WriteLine("2.1 Site publishing feature activated");

                    clientContext.Web.ActivateFeature(publishingWebFeatureId);
                    Assert.IsTrue(clientContext.Web.IsFeatureActive(publishingWebFeatureId));
                    Console.WriteLine("2.2 Web publishing feature activated");

                    // Finally deactivate again
                    clientContext.Web.DeactivateFeature(publishingWebFeatureId);
                    Assert.IsFalse(clientContext.Web.IsFeatureActive(publishingWebFeatureId));
                    Console.WriteLine("2.3 Web publishing feature deactivated");

                    clientContext.Site.DeactivateFeature(publishingSiteFeatureId);
                    Assert.IsFalse(clientContext.Site.IsFeatureActive(publishingSiteFeatureId));
                    Console.WriteLine("2.4 Site publishing feature deactivated");
                }
            }
        }

        [TestMethod()]
        public void ActivateSiteFeatureTest()
        {
            // Test
            clientContext.Site.ActivateFeature(sp2007WorkflowSiteFeatureId);

            Assert.IsTrue(clientContext.Site.IsFeatureActive(sp2007WorkflowSiteFeatureId));

            clientContext.Site.DeactivateFeature(sp2007WorkflowSiteFeatureId);
            
            Assert.IsFalse(clientContext.Site.IsFeatureActive(sp2007WorkflowSiteFeatureId));
        }

        [TestMethod()]
        public void ActivateWebFeatureTest()
        {
            // Test
            clientContext.Web.ActivateFeature(contentOrganizerWebFeatureId);

            Assert.IsTrue(clientContext.Web.IsFeatureActive(contentOrganizerWebFeatureId));

            clientContext.Web.DeactivateFeature(contentOrganizerWebFeatureId);

            Assert.IsFalse(clientContext.Web.IsFeatureActive(contentOrganizerWebFeatureId));
        }

        [TestMethod()]
        public void DeactivateSiteFeatureTest()
        {
            // Setup
            clientContext.Site.ActivateFeature(sp2007WorkflowSiteFeatureId);

            // Test
            clientContext.Site.DeactivateFeature(sp2007WorkflowSiteFeatureId);
            Assert.IsFalse(clientContext.Site.IsFeatureActive(sp2007WorkflowSiteFeatureId));
        }

        [TestMethod()]
        public void DeactivateWebFeatureTest()
        {
            // Setup
            clientContext.Web.ActivateFeature(contentOrganizerWebFeatureId);

            // Test
            clientContext.Web.DeactivateFeature(contentOrganizerWebFeatureId);
            Assert.IsFalse(clientContext.Web.IsFeatureActive(contentOrganizerWebFeatureId));
        }

        [TestMethod()]
        public void IsSiteFeatureActiveTest()
        {
            // Setup
            try
            {
                clientContext.Site.DeactivateFeature(sp2007WorkflowSiteFeatureId);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ignoring exception: {0}", ex.Message);
            }

            // Test
            Assert.IsFalse(clientContext.Site.IsFeatureActive(sp2007WorkflowSiteFeatureId));
        }

        [TestMethod()]
        public void IsWebFeatureActiveTest()
        {
            // Setup
            try
            { 
                clientContext.Web.DeactivateFeature(contentOrganizerWebFeatureId);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ignoring exception: {0}", ex.Message);
            }

            // Test
            Assert.IsFalse(clientContext.Web.IsFeatureActive(contentOrganizerWebFeatureId));
        }
#endregion
    }
}
