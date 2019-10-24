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
    public class CrossSiteTests
    {
        class TestLayout : ILayoutTransformator
        {
            public void Transform(Tuple<Pages.PageLayout, List<WebPartEntity>> pageData)
            {
                throw new NotImplementedException();
            }
        }

        class TestTransformator : IContentTransformator
        {
            public void Transform(List<WebPartEntity> webParts)
            {
                throw new NotImplementedException();
            }
        }


        #region Test initialization
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            //using (var cc = TestCommon.CreateClientContext())
            //{
            //    // Clean all migrated pages before next run
            //    var pages = cc.Web.GetPages("Migrated_");

            //    foreach (var page in pages.ToList())
            //    {
            //        page.DeleteObject();
            //    }

            //    cc.ExecuteQueryRetry();
            //}
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {

        }
        #endregion



        [TestMethod]
        public void CrossSiteTransformTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevTeamSiteUrl")))
                {
                    var pageTransformator = new PageTransformator(sourceClientContext, targetClientContext);
                    pageTransformator.RegisterObserver(new UnitTestLogObserver());

                    var pages = sourceClientContext.Web.GetPages("wpp_with");

                    foreach (var page in pages)
                    {
                        PageTransformationInformation pti = new PageTransformationInformation(page)
                        {
                            // If target page exists, then overwrite it
                            Overwrite = true,

                            // Don't log test runs
                            SkipTelemetry = true,

                            // ModernizationCenter options
                            //ModernizationCenterInformation = new ModernizationCenterInformation()
                            //{
                            //    AddPageAcceptBanner = true
                            //},

                            // Migrated page gets the name of the original page
                            //TargetPageTakesSourcePageName = true,

                            // Give the migrated page a specific prefix, default is Migrated_
                            //TargetPagePrefix = "Yes_",

                            // Configure the page header, empty value means ClientSidePageHeaderType.None
                            //PageHeader = new ClientSidePageHeader(cc, ClientSidePageHeaderType.None, null),

                            // If the page is a home page then replace with stock home page
                            //ReplaceHomePageWithDefaultHomePage = true,

                            // Replace embedded images and iframes with a placeholder and add respective images and video web parts at the bottom of the page
                            // HandleWikiImagesAndVideos = false,

                            // Callout to your custom code to allow for title overriding
                            //PageTitleOverride = titleOverride,

                            // Callout to your custom layout handler
                            //LayoutTransformatorOverride = layoutOverride,

                            // Callout to your custom content transformator...in case you fully want replace the model
                            //ContentTransformatorOverride = contentOverride,
                        };

                        pti.MappingProperties["SummaryLinksToQuickLinks"] = "true";
                        pti.MappingProperties["UseCommunityScriptEditor"] = "true";

                        var result = pageTransformator.Transform(pti);
                    }
                }
            }

            //Assert.Inconclusive(TestCommon.InconclusiveNoAutomatedChecksMessage);

        }

        [TestMethod]
        public void CrossSiteTransform_OverwriteOffTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                //Test Requires a test site
                using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevTeamSiteUrl")))
                {
                    var pageTransformator = new PageTransformator(sourceClientContext, targetClientContext);
                    pageTransformator.RegisterObserver(new UnitTestLogObserver());

                    var pages = sourceClientContext.Web.GetPages("wpp_with"); //Specific page - aim for one file

                    foreach (var page in pages)
                    {
                        PageTransformationInformation pti = new PageTransformationInformation(page)
                        {
                            // If target page exists, then overwrite it
                            Overwrite = false,

                            // Don't log test runs
                            SkipTelemetry = true,

                        };

                        pti.MappingProperties["SummaryLinksToQuickLinks"] = "true";
                        pti.MappingProperties["UseCommunityScriptEditor"] = "true";

                        var result = Assert.ThrowsException<ArgumentException>(() =>
                        {
                            var result1 = pageTransformator.Transform(pti);
                            var result2 = pageTransformator.Transform(pti); //Run twice incase target site didnt have the file in the first place
                        });

                        Assert.IsTrue(result.Message.Contains("Not overwriting - there already exists a page with name"));

                    }
                }
            }
        }


        [TestMethod]
        public void CrossSiteTransform_SameSite_WebPartPageTest()
        {
            using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevTeamSiteUrl")))
            {
                var pageTransformator = new PageTransformator(sourceClientContext);
                pageTransformator.RegisterObserver(new UnitTestLogObserver());

                var pages = sourceClientContext.Web.GetPages("wpp_with");

                foreach (var page in pages)
                {
                    PageTransformationInformation pti = new PageTransformationInformation(page)
                    {
                        // If target page exists, then overwrite it
                        Overwrite = true,

                        // Don't log test runs
                        SkipTelemetry = true,

                        // ModernizationCenter options
                        ModernizationCenterInformation = new ModernizationCenterInformation()
                        {
                            AddPageAcceptBanner = true
                        },

                        // Migrated page gets the name of the original page
                        //TargetPageTakesSourcePageName = true,

                        // Give the migrated page a specific prefix, default is Migrated_
                        TargetPagePrefix = "Converted_",

                        // Configure the page header, empty value means ClientSidePageHeaderType.None
                        //PageHeader = new ClientSidePageHeader(cc, ClientSidePageHeaderType.None, null),

                        // If the page is a home page then replace with stock home page
                        //ReplaceHomePageWithDefaultHomePage = true,

                        // Replace embedded images and iframes with a placeholder and add respective images and video web parts at the bottom of the page
                        HandleWikiImagesAndVideos = false,

                        // Callout to your custom code to allow for title overriding
                        //PageTitleOverride = titleOverride,

                        // Callout to your custom layout handler
                        //LayoutTransformatorOverride = layoutOverride,

                        // Callout to your custom content transformator...in case you fully want replace the model
                        //ContentTransformatorOverride = contentOverride,
                    };

                    pageTransformator.Transform(pti);
                }

            }

            Assert.Inconclusive(TestCommon.InconclusiveNoAutomatedChecksMessage);

        }

        [TestMethod]
        public void CrossSiteTransform_SameSite_OverwriteOffTest()
        {

            //Test Requires a test site
            using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevTeamSiteUrl")))
            {
                var pageTransformator = new PageTransformator(sourceClientContext);
                pageTransformator.RegisterObserver(new UnitTestLogObserver());

                var pages = sourceClientContext.Web.GetPages("wpp_with"); //Specific page - aim for one file

                foreach (var page in pages)
                {
                    PageTransformationInformation pti = new PageTransformationInformation(page)
                    {
                        // If target page exists, then overwrite it
                        Overwrite = false,

                        // Don't log test runs
                        SkipTelemetry = true,

                        TargetPagePrefix = "Converted_",

                    };

                    pti.MappingProperties["SummaryLinksToQuickLinks"] = "true";
                    pti.MappingProperties["UseCommunityScriptEditor"] = "true";

                    var result = Assert.ThrowsException<ArgumentException>(() =>
                    {
                        var result1 = pageTransformator.Transform(pti);
                        var result2 = pageTransformator.Transform(pti); //Run twice incase target site didnt have the file in the first place
                        });

                    Assert.IsTrue(result.Message.Contains("Not overwriting - there already exists a page with name"));

                }
            }

        }


        [TestMethod]
        public void CrossSiteTransform_SameSite_WikiPageTest()
        {
            using (var sourceClientContext = TestCommon.CreateClientContext())
            {
                var pageTransformator = new PageTransformator(sourceClientContext);

                var pages = sourceClientContext.Web.GetPages("wk");

                foreach (var page in pages)
                {
                    PageTransformationInformation pti = new PageTransformationInformation(page)
                    {
                        // If target page exists, then overwrite it
                        Overwrite = true,

                        // Don't log test runs
                        SkipTelemetry = true,

                        // ModernizationCenter options
                        ModernizationCenterInformation = new ModernizationCenterInformation()
                        {
                            AddPageAcceptBanner = true
                        },

                        // Migrated page gets the name of the original page
                        //TargetPageTakesSourcePageName = true,

                        // Give the migrated page a specific prefix, default is Migrated_
                        TargetPagePrefix = "Converted_",

                        // Configure the page header, empty value means ClientSidePageHeaderType.None
                        //PageHeader = new ClientSidePageHeader(cc, ClientSidePageHeaderType.None, null),

                        // If the page is a home page then replace with stock home page
                        //ReplaceHomePageWithDefaultHomePage = true,

                        // Replace embedded images and iframes with a placeholder and add respective images and video web parts at the bottom of the page
                        HandleWikiImagesAndVideos = false,

                        // Callout to your custom code to allow for title overriding
                        //PageTitleOverride = titleOverride,

                        // Callout to your custom layout handler
                        //LayoutTransformatorOverride = layoutOverride,

                        // Callout to your custom content transformator...in case you fully want replace the model
                        //ContentTransformatorOverride = contentOverride,
                    };

                    pageTransformator.Transform(pti);
                }

            }

            Assert.Inconclusive(TestCommon.InconclusiveNoAutomatedChecksMessage);

        }
    }
}
