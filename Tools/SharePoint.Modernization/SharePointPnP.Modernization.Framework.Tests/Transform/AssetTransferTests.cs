﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Transform;
using OfficeDevPnP.Core.Pages;
using SharePointPnP.Modernization.Framework.Pages;
using SharePointPnP.Modernization.Framework.Entities;
using System.Linq;

namespace SharePointPnP.Modernization.Framework.Tests.Transform
{
    [TestClass]
    public class AssetTransferTests
    {

        /// <summary>
        /// This test validates with SharePoint the entire operation
        /// </summary>
        [TestMethod]
        public void AssetTransfer_CopyAssetToTarget_SmallFileTest()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevTeamSiteUrl")))
                {
                    AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, targetClientContext);

                    var targetWebUrl = targetClientContext.Web.EnsureProperty(o => o.ServerRelativeUrl);
                    var sourceWebUrl = sourceClientContext.Web.EnsureProperty(o => o.ServerRelativeUrl);

                    // Very crude test - ensure the site is setup for this ahead of the test
                    var sourceFileServerRelativeUrl = $"{sourceWebUrl}/SiteImages/extra8_500x500.jpg";
                    var targetLocation = $"{targetWebUrl}/Shared%20Documents"; //Shared Documents for example, Site Assets may not exist on vanilla sites

                    assetTransfer.CopyAssetToTargetLocation(sourceFileServerRelativeUrl, targetLocation);
                }
            }

            Assert.Inconclusive(TestCommon.InconclusiveNoAutomatedChecksMessage);
        }

        /// <summary>
        /// This test validates with SharePoint the entire operation
        /// </summary>
        [TestMethod]
        public void AssetTransfer_CopyAssetToTarget_LargeFileTest()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevTeamSiteUrl")))
                {
                    AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, targetClientContext);

                    // Very crude test - ensure the site is setup for this ahead of the test
                    // Note this file is not included in this project assets due to its licensing. Pls find a > 3MB file to use as a test.
                    var targetWebUrl = targetClientContext.Web.EnsureProperty(o => o.ServerRelativeUrl);
                    var sourceWebUrl = sourceClientContext.Web.EnsureProperty(o => o.ServerRelativeUrl);

                    var sourceFileServerRelativeUrl = $"{sourceWebUrl}/SiteImages/bigstock-Html-Web-Code-57446159.jpg";
                    var targetLocation = $"{targetWebUrl}/Shared%20Documents"; //Shared Documents for example, Site Assets may not exist on vanilla sites

                    assetTransfer.CopyAssetToTargetLocation(sourceFileServerRelativeUrl, targetLocation);
                }
            }

            Assert.Inconclusive(TestCommon.InconclusiveNoAutomatedChecksMessage);
        }

        /// <summary>
        /// This test validates with SharePoint the entire operation
        /// </summary>
        [TestMethod]
        public void AssetTransfer_CopyAssetToTarget_PagesWithImageWebPartTest()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevTeamSiteUrl")))
                {
                    var pageTransformator = new PageTransformator(sourceClientContext, targetClientContext);

                    var pages = sourceClientContext.Web.GetPages("WPP_Image-Asset");

                    foreach (var page in pages)
                    {
                        PageTransformationInformation pti = new PageTransformationInformation(page)
                        {
                            // If target page exists, then overwrite it
                            Overwrite = true,

                            // Don't log test runs
                            SkipTelemetry = true,

                            // Replace embedded images and iframes with a placeholder and add respective images and video web parts at the bottom of the page
                            //HandleWikiImagesAndVideos = false,
                        };

                        pageTransformator.Transform(pti);
                    }
                }
            }

            Assert.Inconclusive(TestCommon.InconclusiveNoAutomatedChecksMessage);
        }

        /// <summary>
        /// This test validates with SharePoint the entire operation
        /// </summary>
        [TestMethod]
        public void AssetTransfer_ValidateSupportedLocation_AspxRejectTest()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext())
                {
                    // Needs valid client contexts as they are part of the checks.
                    AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, targetClientContext);

                    var webUrl = sourceClientContext.Web.EnsureProperty(o => o.ServerRelativeUrl);
                    var result = assetTransfer.ValidateAssetInSupportedLocation($"{webUrl}/siteassets/wrongfile.aspx");
                    var expected = false;

                    Assert.AreEqual(expected, result);
                }
            }
        }

        /// <summary>
        /// This test validates with SharePoint the entire operation
        /// </summary>
        [TestMethod]
        public void AssetTransfer_ValidateSupportedLocation_OtherTenantRejectTest()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext())
                {
                    // Needs valid client contexts as they are part of the checks.
                    AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, targetClientContext);
                    var result = assetTransfer.ValidateAssetInSupportedLocation("https://faketenant.sharepoint.com/sites/fakesitecollection/images/afakeimage.jpg");
                    var expected = false;

                    Assert.AreEqual(expected, result);
                }
            }
        }

        /// <summary>
        /// This test validates with SharePoint the entire operation
        /// </summary>
        [TestMethod]
        public void AssetTransfer_ValidateSupportedLocation_OtherSiteCollectionRelativeRejectTest()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext())
                {
                    // Needs valid client contexts as they are part of the checks.
                    AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, targetClientContext);
                    var result = assetTransfer.ValidateAssetInSupportedLocation($"/sites/fakesitecollection/images/afakeimage.jpg");
                    var expected = false;

                    Assert.AreEqual(expected, result);
                }
            }
        }

        [TestMethod]
        public void AssetTransfer_ValidateSupportedLocation_ImageOnSubsiteAcceptTest()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext())
                {
                    // Needs valid client contexts as they are part of the checks.
                    AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, targetClientContext);

                    var webUrl = sourceClientContext.Web.EnsureProperty(o => o.ServerRelativeUrl);
                    
                    var result = assetTransfer.ValidateAssetInSupportedLocation($"{webUrl}/subsite/siteassets/afakeimage.jpg");
                    var expected = true;

                    Assert.AreEqual(expected, result);
                }
            }
        }

        /// <summary>
        /// This test validates with SharePoint the entire operation
        /// </summary>
        [TestMethod]
        public void AssetTransfer_ValiateSupportedLocation_SameCtxRejectTest()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext())
            {
                using (var sourceClientContext = TestCommon.CreateClientContext())
                {
                    // Needs valid client contexts as they are part of the checks.
                    AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, targetClientContext);
                    var webUrl = sourceClientContext.Web.EnsureProperty(o => o.ServerRelativeUrl);
                    var result = assetTransfer.ValidateAssetInSupportedLocation($"{webUrl}/siteassets/rightfile.jpg");
                    var expected = false;

                    Assert.AreEqual(expected, result);
                }
            }
        }

        /// <summary>
        /// This test validates with SharePoint the entire operation
        /// </summary>
        [TestMethod]
        public void AssetTransfer_ValiateSupportedLocation_AcceptTest()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext())
                {
                    // Needs valid client contexts as they are part of the checks.
                    AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, targetClientContext);

                    var webUrl = sourceClientContext.Web.EnsureProperty(o => o.ServerRelativeUrl);
                    var result = assetTransfer.ValidateAssetInSupportedLocation($"{webUrl}/siteassets/rightfile.jpg");
                    var expected = true;

                    Assert.AreEqual(expected, result);
                }
            }
        }

        /// <summary>
        /// This test validates with SharePoint the entire operation
        /// </summary>
        [TestMethod]
        public void AssetTransfer_CopyAssetToTarget_WithCacheSameFileTest()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevTeamSiteUrl")))
                {
                    var pageTransformator = new PageTransformator(sourceClientContext, targetClientContext);

                    var pages = sourceClientContext.Web.GetPages("WPP_Image-Asset-MultipleImages-Test");

                    foreach (var page in pages)
                    {
                        PageTransformationInformation pti = new PageTransformationInformation(page)
                        {
                            // If target page exists, then overwrite it
                            Overwrite = true,

                            // Don't log test runs
                            SkipTelemetry = true,

                            // Replace embedded images and iframes with a placeholder and add respective images and video web parts at the bottom of the page
                            //HandleWikiImagesAndVideos = false,
                        };

                        pageTransformator.Transform(pti);
                    }
                }
            }

            Assert.Inconclusive(TestCommon.InconclusiveNoAutomatedChecksMessage);
        }

        /// <summary>
        /// This test validates with SharePoint the entire operation
        /// </summary>
        [TestMethod]
        public void AssetTransfer_CopyAssetToTarget_WithCacheMultipleFileTest()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevTeamSiteUrl")))
                {
                    var pageTransformator = new PageTransformator(sourceClientContext, targetClientContext);

                    var pages = sourceClientContext.Web.GetPages("WPP_Image-Asset");

                    foreach (var page in pages)
                    {
                        PageTransformationInformation pti = new PageTransformationInformation(page)
                        {
                            // If target page exists, then overwrite it
                            Overwrite = true,

                            // Don't log test runs
                            SkipTelemetry = true,

                            // Replace embedded images and iframes with a placeholder and add respective images and video web parts at the bottom of the page
                            //HandleWikiImagesAndVideos = false,
                        };

                        pageTransformator.Transform(pti);
                    }
                }
            }

            Assert.Inconclusive(TestCommon.InconclusiveNoAutomatedChecksMessage);
        }

        /// <summary>
        /// This test validates with SharePoint the entire operation
        /// </summary>
        [TestMethod]
        public void AssetTransfer_CopyAssetToTarget_WithFullTransformTest()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevTeamSiteUrl")))
                {
                    var pageTransformator = new PageTransformator(sourceClientContext, targetClientContext);

                    var pages = sourceClientContext.Web.GetPages("p");

                    foreach (var page in pages)
                    {
                        PageTransformationInformation pti = new PageTransformationInformation(page)
                        {
                            // If target page exists, then overwrite it
                            Overwrite = true,

                            // Don't log test runs
                            SkipTelemetry = true,

                            // Replace embedded images and iframes with a placeholder and add respective images and video web parts at the bottom of the page
                            //HandleWikiImagesAndVideos = false,
                        };

                        pageTransformator.Transform(pti);
                    }
                }
            }

            Assert.Inconclusive(TestCommon.InconclusiveNoAutomatedChecksMessage);
        }

        /// <summary>
        /// This test validates with SharePoint the entire operation
        /// </summary>
        [TestMethod]
        public void AssetTransfer_EnsureSiteAssets_Test()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext())
                {
                    AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, targetClientContext);
                    assetTransfer.EnsureSiteAssetsLibrary();

                    // Validate the site assets works
                    var siteAssetsExist = targetClientContext.Web.ListExists("Site Assets");
                    Assert.IsTrue(siteAssetsExist);
                    // Clean up the test target site

                }
            }
        }

        /// <summary>
        /// This test validates with SharePoint the entire operation
        /// </summary>
        [TestMethod]
        public void AssetTransfer_EnsureFolderLocaton_Test()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext())
                {
                    // Needs valid client contexts as they are part of the checks.
                    AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, targetClientContext);
                    var result = assetTransfer.EnsureDestination("WPP_Image-Asset-Test.aspx");

                    var webUrl = targetClientContext.Web.EnsureProperty(o => o.ServerRelativeUrl);
                    var expected = $"{webUrl}/SiteAssets/SitePages/WPP_Image-Asset-Test";

                    Assert.AreEqual(expected, result);
                }
            }
        }

        /// <summary>
        /// This test validates with SharePoint the entire operation
        /// </summary>
        [TestMethod]
        public void AssetTransfer_TransferAsset_Test()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevTeamSiteUrl")))
                {
                    // Needs valid client contexts as they are part of the checks.
                    AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, targetClientContext);
                    
                    var targetWebUrl = targetClientContext.Web.EnsureProperty(o => o.ServerRelativeUrl);
                    var sourceWebUrl = sourceClientContext.Web.EnsureProperty(o => o.ServerRelativeUrl);

                    // Very crude test - ensure the site is setup for this ahead of the test
                    var sourceFileServerRelativeUrl = $"{sourceWebUrl}/SiteImages/extra8_500x500.jpg";
                    var targetLocation = $"{targetWebUrl}/Shared%20Documents"; //Shared Documents for example, Site Assets may not exist on vanilla sites

                    var result = assetTransfer.TransferAsset(sourceFileServerRelativeUrl, "This is a unit test page.aspx"); //Page shouldnt need to exist at this point
                    var expected = $"{targetWebUrl}/SiteAssets/SitePages/This-is-a-unit-test-page/extra8_500x500.jpg";

                    Assert.AreEqual(expected, result);
                }
            }
        }

        /// <summary>
        /// This test validates with SharePoint the entire operation
        /// </summary>
        [TestMethod]
        public void AssetTransfer_TransferAsset_OnPrem_Test()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateOnPremisesClientContext())
                {
                    // Needs valid client contexts as they are part of the checks.
                    AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, targetClientContext);

                    var targetWebUrl = targetClientContext.Web.EnsureProperty(o => o.ServerRelativeUrl);
                    var sourceWebUrl = sourceClientContext.Web.EnsureProperty(o => o.ServerRelativeUrl);

                    var sourceFile = "shutterstock_113327035.jpg";

                    // Very crude test - ensure the site is setup for this ahead of the test
                    var sourceFileServerRelativeUrl = $"{sourceWebUrl}PublishingImages/{sourceFile}";
                    var targetLocation = $"{targetWebUrl}/Shared%20Documents"; //Shared Documents for example, Site Assets may not exist on vanilla sites

                    var result = assetTransfer.TransferAsset(sourceFileServerRelativeUrl, "This is a unit test page.aspx"); //Page shouldnt need to exist at this point
                    var expected = $"{targetWebUrl}/SiteAssets/SitePages/This-is-a-unit-test-page/{sourceFile}";

                    Assert.AreEqual(expected, result);
                }
            }
        }

        /// <summary>
        /// This test validates with SharePoint the entire operation
        /// </summary>
        [TestMethod]
        public void AssetTransfer_TransferAsset_OnPrem_Large_Test()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateOnPremisesClientContext())
                {
                    // Needs valid client contexts as they are part of the checks.
                    AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, targetClientContext);

                    var targetWebUrl = targetClientContext.Web.EnsureProperty(o => o.ServerRelativeUrl);
                    var sourceWebUrl = sourceClientContext.Web.EnsureProperty(o => o.ServerRelativeUrl);

                    var sourceFile = "shutterstock_171790418.jpg";

                    // Very crude test - ensure the site is setup for this ahead of the test
                    var sourceFileServerRelativeUrl = $"{sourceWebUrl}PublishingImages/{sourceFile}";
                    var targetLocation = $"{targetWebUrl}/Shared%20Documents"; //Shared Documents for example, Site Assets may not exist on vanilla sites

                    var result = assetTransfer.TransferAsset(sourceFileServerRelativeUrl, "This is a unit test page.aspx"); //Page shouldnt need to exist at this point
                    var expected = $"{targetWebUrl}/SiteAssets/SitePages/This-is-a-unit-test-page/{sourceFile}";

                    Assert.AreEqual(expected, result);
                }
            }
        }

        [TestMethod]
        public void AssetTransfer_TransferAsset_ContextSiteImageOnSubSiteAcceptTest()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint
            // Scenario here is when the source context is a subsite and the image asset is on the Site Collection

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPORootDevSiteUrl")))
                {
                    // Needs valid client contexts as they are part of the checks.
                    AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, targetClientContext);

                    var siteUrl = sourceClientContext.Site.EnsureProperty(o => o.ServerRelativeUrl);
                    var target = $"{TestCommon.AppSetting("SPOTargetSiteUrl")}/SiteAssets/";

                    var sourceImageLocation = $"{siteUrl}/en/PublishingImages/extra.jpg";

                    var result = assetTransfer.CopyAssetToTargetLocation(sourceImageLocation, target);
                    Assert.IsTrue(!string.IsNullOrEmpty(result));
                }
            }
        }

        [TestMethod]
        public void AssetTransfer_TransferAsset_EnsureContextRootSiteTest()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint
            // Scenario here is when the source context is a subsite and the image asset is on the Site Collection

            AssetTransferOfAsset("SiteCollectionImages/extra.jpg");

        }

        [TestMethod]
        public void AssetTransfer_TransferAsset_EnsureContextSubSiteTest()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint
            // Scenario here is when the source context is a subsite and the image asset is on the Site Collection

            AssetTransferOfAsset("en/PublishingImages/extra.jpg");

        }

       
        [TestMethod]
        public void AssetTransfer_TransferAsset_EnsureContextDeepTest()
        {
            //Note: This is more of a system test rather than unit given its dependency on SharePoint
            // Scenario here is when the source context is a subsite and the image asset is on the Site Collection

            AssetTransferOfAsset("en/subsite84/PublishingImages/Folder/extra.jpg");

        }

        /// <summary>
        /// Note: this test will only work if the files are present at the parts of the sharepoint site
        /// and the source site originates from a sub site
        /// </summary>
        /// <param name="url"></param>
        private void AssetTransferOfAsset(string url)
        {
            if (string.IsNullOrEmpty(url))
            {
                Assert.Inconclusive("You havent provided a URL");
            }

            url = url.TrimStart('/'); //Remove slash at the beginning

            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevSiteUrl")))
                {
                    // Needs valid client contexts as they are part of the checks.
                    AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, targetClientContext);
                    assetTransfer.RegisterObserver(new UnitTestLogObserver());

                    var siteUrl = sourceClientContext.Site.EnsureProperty(o => o.ServerRelativeUrl);
                    var target = $"{TestCommon.AppSetting("SPOTargetSiteUrl")}/SiteAssets/";

                    var sourceImageLocation = $"{siteUrl}/{url}";

                    assetTransfer.EnsureAssetContextIfRequired(sourceClientContext, sourceImageLocation);

                    var result = assetTransfer.CopyAssetToTargetLocation(sourceImageLocation, target);

                    Assert.IsTrue(!string.IsNullOrEmpty(result));
                }
            }
        }


    }
}
