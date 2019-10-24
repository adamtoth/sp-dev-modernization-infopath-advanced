﻿using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SharePointPnP.Modernization.Framework.Pages;
using SharePointPnP.Modernization.Framework.Publishing;
using SharePointPnP.Modernization.Framework.Telemetry;
using SharePointPnP.Modernization.Framework.Telemetry.Observers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SharePointPnP.Modernization.Framework.Tests.Transform.OnPremises
{
    [TestClass]
    public class OnPremisesPublishingPageTests
    {
        [TestMethod]
        public void BasicOnPremPublishingPageTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                using (var sourceClientContext = TestCommon.CreateOnPremisesClientContext())
                {
                    //"C:\github\sp-dev-modernization\Tools\SharePoint.Modernization\SharePointPnP.Modernization.Framework.Tests\Transform\Publishing\custompagelayoutmapping.xml"
                    //"C:\temp\onprem-mapping-all-test.xml.xml"
                    var pageTransformator = new PublishingPageTransformator(sourceClientContext, targetClientContext, @"C:\temp\onprem-mapping-all-test.xml");
                    pageTransformator.RegisterObserver(new MarkdownObserver(folder: "c:\\temp", includeVerbose: true));
                    pageTransformator.RegisterObserver(new UnitTestLogObserver());

                    //var pages = sourceClientContext.Web.GetPagesFromList("Pages", "Article-2010-Custom");
                    //var pages = sourceClientContext.Web.GetPagesFromList("Pages", "ArticlePage-2010-Multiple");
                    //var pages = sourceClientContext.Web.GetPagesFromList("Pages", "Article-2010-Custom-Test3");
                    //var pages = sourceClientContext.Web.GetPagesFromList("Pages", folder:"News");
                    //var pages = sourceClientContext.Web.GetPagesFromList("Pages", "Welcome-2013Legacy");
                    var pages = sourceClientContext.Web.GetPagesFromList("Pages", "Welcome-SP2013");

                    pages.FailTestIfZero();

                    foreach (var page in pages)
                    {
                        PublishingPageTransformationInformation pti = new PublishingPageTransformationInformation(page)
                        {
                            // If target page exists, then overwrite it
                            Overwrite = true,

                            // Don't log test runs
                            SkipTelemetry = true,

                            //Permissions are unlikely to work given cross domain
                            KeepPageSpecificPermissions = false,

                            //RemoveEmptySectionsAndColumns = false,

                            // Configure the page header, empty value means ClientSidePageHeaderType.None
                            //PageHeader = new ClientSidePageHeader(cc, ClientSidePageHeaderType.None, null),

                            // Replace embedded images and iframes with a placeholder and add respective images and video web parts at the bottom of the page
                            // HandleWikiImagesAndVideos = false,

                            // Callout to your custom code to allow for title overriding
                            //PageTitleOverride = titleOverride,

                            // Callout to your custom layout handler
                            //LayoutTransformatorOverride = layoutOverride,

                            // Callout to your custom content transformator...in case you fully want replace the model
                            //ContentTransformatorOverride = contentOverride,
                            //SkipUrlRewrite = true
                        };

                        Console.WriteLine("SharePoint Version: {0}", pti.SourceVersion);

                        pti.MappingProperties["SummaryLinksToQuickLinks"] = "true";
                        pti.MappingProperties["UseCommunityScriptEditor"] = "true";

                        var result = pageTransformator.Transform(pti);
                    }

                    pageTransformator.FlushObservers();
                    
                }
            }
        }

        [TestMethod]
        public void OnPremPageLayout_AnalyzeByPages_Test()
        {
            using (var context = TestCommon.CreateOnPremisesClientContext())
            {
                var pages = context.Web.GetPagesFromList("Pages");
                var analyzer = new PageLayoutAnalyser(context);
                int errorCount = 0;
                foreach (var page in pages)
                {
                    try
                    {
                        analyzer.AnalysePageLayoutFromPublishingPage(page);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error {0} {1}", ex.Message, ex.StackTrace);
                        errorCount++;
                    }
                }

                Console.WriteLine("Error Count {0}", errorCount);
                Assert.IsTrue((errorCount == 0));
                analyzer.GenerateMappingFile("c:\\temp", "onprem-mapping-test.xml");
            }
        }

        [TestMethod]
        public void OnPremPageLayout_AnalyseAll_Test()
        {
            using (var context = TestCommon.CreateOnPremisesClientContext())
            {

                var analyzer = new PageLayoutAnalyser(context);
                analyzer.AnalyseAll();

                analyzer.GenerateMappingFile("c:\\temp", "onprem-mapping-all-test.xml");
            }
        }

        [TestMethod]
        public void BasePage_ExtractWebPartDocumentViaWebServicesFromPageTest()
        {
            string url = "http://portal2010/pages/article-2010-custom.aspx";
            //string url = "/pages/article-2010-custom.aspx";

            using (var context = TestCommon.CreateOnPremisesClientContext())
            {

                var pages = context.Web.GetPagesFromList("Pages", "Article-2010-Custom");

                foreach (var page in pages)
                {
                    page.EnsureProperty(p => p.File);

                    List<string> search = new List<string>()
                    {
                        "WebPartZone"
                    };

                    //Should be one
                    TestBasePage testBase = new TestBasePage(page, page.File, null, null);
                    var result = testBase.ExtractWebPartDocumentViaWebServicesFromPage(url);

                    Assert.IsTrue(result.Item1.Length > 0);
                    Assert.IsTrue(result.Item2.Length > 0);

                    break;

                }
            }

        }

        [TestMethod]
        public void BasePage_LoadWebPartDocumentViaWebServicesTest()
        {
            //string url = "http://portal2010/pages/article-2010-custom.aspx";
            //string url = "/pages/article-2010-custom.aspx";
            //string url = "/pages/article-2010-custom.aspx";
            string url = "/pages/welcome-2013.aspx";

            using (var context = TestCommon.CreateOnPremisesClientContext())
            {

                var pages = context.Web.GetPagesFromList("Pages", "Article-2010-Custom-Test2");

                foreach (var page in pages)
                {
                    page.EnsureProperty(p => p.File);

                    List<string> search = new List<string>()
                    {
                        "WebPartZone"
                    };

                    //Should be one
                    TestBasePage testBase = new TestBasePage(page, page.File, null, null);
                    var result = testBase.LoadPublishingPageFromWebServices(url);

                    Assert.IsTrue(result.Count > 0);

                }
            }

        }

       


        [TestMethod]
        public void BasePage_ExportWebPartByWorkaround()
        {
            //string url = "http://portal2010/pages/article-2010-custom.aspx";
            string url = "/pages/article-2010-custom-test2.aspx";
            //string url = "/pages/article-2010-custom.aspx";

            using (var context = TestCommon.CreateOnPremisesClientContext())
            {

                var pages = context.Web.GetPagesFromList("Pages", "Article-2010-Custom-Test2");

                foreach (var page in pages)
                {
                    page.EnsureProperty(p => p.File);

                    //Should be one
                    TestBasePage testBase = new TestBasePage(page, page.File, null, null);
                    var webPartEntities = testBase.LoadPublishingPageFromWebServices(url);

                    foreach (var webPart in webPartEntities)
                    {
                        var result = testBase.ExportWebPartXmlWorkaround(url, webPart.Id.ToString());

                        Assert.IsTrue(!string.IsNullOrEmpty(result));

                    }

                }
            }

        }
    

}


public class TestBasePage : BasePage
{
    public TestBasePage(ListItem item, File file, PageTransformation pt, IList<ILogObserver> logObservers) : base(item, file, pt, logObservers)
    {

    }
}

}
