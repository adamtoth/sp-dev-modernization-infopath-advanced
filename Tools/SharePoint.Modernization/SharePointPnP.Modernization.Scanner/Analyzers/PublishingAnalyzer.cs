﻿using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Entities;
using SharePoint.Modernization.Scanner.Results;
using SharePoint.Scanning.Framework;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml.XPath;

namespace SharePoint.Modernization.Scanner.Analyzers
{
    public class PublishingAnalyzer : BaseAnalyzer
    {
        const string AvailablePageLayouts = "__PageLayouts";
        const string DefaultPageLayout = "__DefaultPageLayout";
        const string AvailableWebTemplates = "__WebTemplates";
        const string InheritWebTemplates = "__InheritWebTemplates";
        const string WebNavigationSettings = "_webnavigationsettings";
        const string FileRefField = "FileRef";
        const string FileLeafRefField = "FileLeafRef";
        const string PublishingPageLayoutField = "PublishingPageLayout";

        // Queries
        const string CAMLQueryByExtension = @"
                <View Scope='RecursiveAll'>
                  <Query>
                    <Where>
                      <Contains>
                        <FieldRef Name='File_x0020_Type'/>
                        <Value Type='text'>aspx</Value>
                      </Contains>
                    </Where>
                  </Query>
                  <ViewFields>
                    <FieldRef Name='ContentTypeId' />
                    <FieldRef Name='FileRef' />
                    <FieldRef Name='FileLeafRef' />
                    <FieldRef Name='File_x0020_Type' />
                    <FieldRef Name='Editor' />
                    <FieldRef Name='Modified' />
                    <FieldRef Name='PublishingPageLayout' />
                    <FieldRef Name='Audience' />
                    <FieldRef Name='PublishingRollupImage' />
                  </ViewFields>  
                </View>";

        // Stores page customization information
        public Dictionary<string, CustomizedPageStatus> MasterPageGalleryCustomization = null;

        private WebScanResult webScanResult;
        private SiteScanResult siteScanResult;

        #region Construction
        /// <summary>
        /// Publishing analyzer construction
        /// </summary>
        /// <param name="url">Url of the web to be analyzed</param>
        /// <param name="siteColUrl">Url of the site collection hosting this web</param>
        public PublishingAnalyzer(string url, string siteColUrl, ModernizationScanJob scanJob, WebScanResult webScanResult, SiteScanResult siteScanResult) : base(url, siteColUrl, scanJob)
        {
            this.webScanResult = webScanResult;
            this.siteScanResult = siteScanResult;
        }
        #endregion

        public override TimeSpan Analyze(ClientContext cc)
        {
            try
            {
                base.Analyze(cc);

                // Only scan when it's a valid publishing portal
                var pageCount = ContinueScanning(cc);
                if (pageCount > 0 || pageCount == -1)
                {
                    try
                    {
                        PublishingWebScanResult scanResult = new PublishingWebScanResult()
                        {
                            SiteColUrl = this.SiteCollectionUrl,
                            SiteURL = this.SiteUrl,
                            WebRelativeUrl = this.SiteUrl.Replace(this.SiteCollectionUrl, ""),
                            WebTemplate = this.webScanResult.WebTemplate,
                            BrokenPermissionInheritance = this.webScanResult.BrokenPermissionInheritance,
                            PageCount = pageCount == -1 ? 0 : pageCount,
                            SiteMasterPage = this.webScanResult.CustomMasterPage,
                            SystemMasterPage = this.webScanResult.MasterPage,
                            AlternateCSS = this.webScanResult.AlternateCSS,
                            Admins = this.siteScanResult.Admins,
                            Owners = this.webScanResult.Owners,
                            UserCustomActions = new List<UserCustomActionResult>()
                        };

                        // User custom actions will play a role in complexity calculation
                        if (this.siteScanResult.SiteUserCustomActions != null && this.siteScanResult.SiteUserCustomActions.Count > 0)
                        {
                            scanResult.UserCustomActions.AddRange(this.siteScanResult.SiteUserCustomActions);
                        }
                        if (this.webScanResult.WebUserCustomActions != null && this.webScanResult.WebUserCustomActions.Count > 0)
                        {
                            scanResult.UserCustomActions.AddRange(this.webScanResult.WebUserCustomActions);
                        }

                        Web web = cc.Web;

                        // Load additional web properties
                        web.EnsureProperty(p => p.Language);
                        scanResult.Language = web.Language;

                        // PageLayouts handling
                        var availablePageLayouts = GetPropertyBagValue<string>(web, AvailablePageLayouts, "");
                        var defaultPageLayout = GetPropertyBagValue<string>(web, DefaultPageLayout, "");

                        if (string.IsNullOrEmpty(availablePageLayouts))
                        {
                            scanResult.PageLayoutsConfiguration = "Any";
                        }
                        else if (availablePageLayouts.Equals("__inherit", StringComparison.InvariantCultureIgnoreCase))
                        {
                            scanResult.PageLayoutsConfiguration = "Inherit from parent";
                        }
                        else
                        {
                            scanResult.PageLayoutsConfiguration = "Defined list";

                            try
                            {
                                availablePageLayouts = SanitizeXmlString(availablePageLayouts);

                                // Fill the defined list
                                var element = XElement.Parse(availablePageLayouts);
                                var nodes = element.Descendants("layout");
                                if (nodes != null && nodes.Count() > 0)
                                {
                                    string allowedPageLayouts = "";

                                    foreach (var node in nodes)
                                    {
                                        allowedPageLayouts = allowedPageLayouts + node.Attribute("url").Value.Replace("_catalogs/masterpage/", "") + ",";
                                    }

                                    allowedPageLayouts = allowedPageLayouts.TrimEnd(new char[] { ',' });

                                    scanResult.AllowedPageLayouts = allowedPageLayouts;
                                }
                            }
                            catch(Exception ex)
                            {
                                scanResult.AllowedPageLayouts = "error_retrieving_pagelayouts";
                            }
                        }

                        if (!string.IsNullOrEmpty(defaultPageLayout))
                        {
                            if (defaultPageLayout.Equals("__inherit", StringComparison.InvariantCultureIgnoreCase))
                            {
                                scanResult.DefaultPageLayout = "Inherit from parent";
                            }
                            else
                            {
                                try
                                {
                                    defaultPageLayout = SanitizeXmlString(defaultPageLayout);
                                    var element = XElement.Parse(defaultPageLayout);
                                    scanResult.DefaultPageLayout = element.Attribute("url").Value.Replace("_catalogs/masterpage/", "");
                                }
                                catch (Exception ex)
                                {
                                    scanResult.DefaultPageLayout = "error_retrieving_defaultpagelayout";
                                }
                            }
                        }

                        // Navigation
                        var navigationSettings = web.GetNavigationSettings();
                        if (navigationSettings != null)
                        {
                            if (navigationSettings.GlobalNavigation.ManagedNavigation)
                            {
                                scanResult.GlobalNavigationType = "Managed";
                            }
                            else
                            {
                                scanResult.GlobalNavigationType = "Structural";
                                scanResult.GlobalStructuralNavigationMaxCount = navigationSettings.GlobalNavigation.MaxDynamicItems;
                                scanResult.GlobalStructuralNavigationShowPages = navigationSettings.GlobalNavigation.ShowPages;
                                scanResult.GlobalStructuralNavigationShowSiblings = navigationSettings.GlobalNavigation.ShowSiblings;
                                scanResult.GlobalStructuralNavigationShowSubSites = navigationSettings.GlobalNavigation.ShowSubsites;
                            }

                            if (navigationSettings.CurrentNavigation.ManagedNavigation)
                            {
                                scanResult.CurrentNavigationType = "Managed";
                            }
                            else
                            {
                                scanResult.CurrentNavigationType = "Structural";
                                scanResult.CurrentStructuralNavigationMaxCount = navigationSettings.CurrentNavigation.MaxDynamicItems;
                                scanResult.CurrentStructuralNavigationShowPages = navigationSettings.CurrentNavigation.ShowPages;
                                scanResult.CurrentStructuralNavigationShowSiblings = navigationSettings.CurrentNavigation.ShowSiblings;
                                scanResult.CurrentStructuralNavigationShowSubSites = navigationSettings.CurrentNavigation.ShowSubsites;
                            }

                            if (navigationSettings.GlobalNavigation.ManagedNavigation || navigationSettings.CurrentNavigation.ManagedNavigation)
                            {
                                scanResult.ManagedNavigationAddNewPages = navigationSettings.AddNewPagesToNavigation;
                                scanResult.ManagedNavigationCreateFriendlyUrls = navigationSettings.CreateFriendlyUrlsForNewPages;

                                // get information about the managed nav term set configuration
                                var managedNavXml = GetPropertyBagValue<string>(web, WebNavigationSettings, "");

                                if (!string.IsNullOrEmpty(managedNavXml))
                                {
                                    var managedNavSettings = XElement.Parse(managedNavXml);
                                    IEnumerable<XElement> navNodes = managedNavSettings.XPathSelectElements("./SiteMapProviderSettings/TaxonomySiteMapProviderSettings");
                                    foreach (var node in navNodes)
                                    {
                                        if (node.Attribute("Name").Value.Equals("CurrentNavigationTaxonomyProvider", StringComparison.InvariantCulture))
                                        {
                                            if (node.Attribute("TermSetId") != null)
                                            {
                                                scanResult.CurrentManagedNavigationTermSetId = node.Attribute("TermSetId").Value;
                                            }
                                            else if (node.Attribute("UseParentSiteMap") != null)
                                            {
                                                scanResult.CurrentManagedNavigationTermSetId = "Inherit from parent";
                                            }
                                        }
                                        else if (node.Attribute("Name").Value.Equals("GlobalNavigationTaxonomyProvider", StringComparison.InvariantCulture))
                                        {
                                            if (node.Attribute("TermSetId") != null)
                                            {
                                                scanResult.GlobalManagedNavigationTermSetId = node.Attribute("TermSetId").Value;
                                            }
                                            else if (node.Attribute("UseParentSiteMap") != null)
                                            {
                                                scanResult.GlobalManagedNavigationTermSetId = "Inherit from parent";
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // Pages library
                        List pagesLibrary = null;
                        var lists = web.GetListsToScan();
                        if (lists != null)
                        {
                            pagesLibrary = lists.Where(p => p.BaseTemplate == 850).FirstOrDefault();
                            if (pagesLibrary != null)
                            {
                                pagesLibrary.EnsureProperties(p => p.EnableModeration, p => p.EnableVersioning, p => p.EnableMinorVersions, p => p.EventReceivers, p => p.Fields, p => p.DefaultContentApprovalWorkflowId);
                                scanResult.LibraryEnableModeration = pagesLibrary.EnableModeration;
                                scanResult.LibraryEnableVersioning = pagesLibrary.EnableVersioning;
                                scanResult.LibraryEnableMinorVersions = pagesLibrary.EnableMinorVersions;
                                scanResult.LibraryItemScheduling = pagesLibrary.ItemSchedulingEnabled();
                                scanResult.LibraryApprovalWorkflowDefined = pagesLibrary.DefaultContentApprovalWorkflowId != Guid.Empty;
                            }
                        }

                        // Variations
                        if (scanResult.Level == 0)
                        {
                            var variationLabels = GetVariationLabels(cc);

                            string labels = "";
                            string sourceLabel = "";
                            foreach (var label in variationLabels)
                            {
                                labels = labels + $"{label.Title} ({label.Language}),";

                                if (label.IsSource)
                                {
                                    sourceLabel = label.Title;
                                }

                            }

                            scanResult.VariationLabels = labels.TrimEnd(new char[] { ',' }); ;
                            scanResult.VariationSourceLabel = sourceLabel;
                        }

                        // Scan pages inside the pages library
                        if (pagesLibrary != null && Options.IncludePublishingWithPages(this.ScanJob.Mode))
                        {
                            CamlQuery query = new CamlQuery
                            {
                                ViewXml = CAMLQueryByExtension,
                            };

                            var pages = pagesLibrary.GetItems(query);

                            // Load additional page related information
                            IEnumerable<ListItem> enumerable = web.Context.LoadQuery(pages.IncludeWithDefaultProperties((ListItem item) => item.ContentType));
                            web.Context.ExecuteQueryRetry();

                            if (enumerable.FirstOrDefault() != null)
                            {
                                foreach (var page in enumerable)
                                {
                                    string pageUrl = null;
                                    try
                                    {
                                        if (page.FieldValues.ContainsKey(FileRefField) && !String.IsNullOrEmpty(page[FileRefField].ToString()))
                                        {
                                            pageUrl = page[FileRefField].ToString();
                                        }
                                        else
                                        {
                                            //skip page
                                            continue;
                                        }

                                        // Basic information about the page
                                        PublishingPageScanResult pageScanResult = new PublishingPageScanResult()
                                        {
                                            SiteColUrl = this.SiteCollectionUrl,
                                            SiteURL = this.SiteUrl,
                                            WebRelativeUrl = scanResult.WebRelativeUrl,
                                            PageRelativeUrl = scanResult.WebRelativeUrl.Length > 0 ? pageUrl.Replace(scanResult.WebRelativeUrl, "") : pageUrl,
                                        };

                                        // Page name
                                        if (page.FieldValues.ContainsKey(FileLeafRefField) && !String.IsNullOrEmpty(page[FileLeafRefField].ToString()))
                                        {
                                            pageScanResult.PageName = page[FileLeafRefField].ToString();
                                        }

                                        // Get page change information
                                        pageScanResult.ModifiedAt = page.LastModifiedDateTime();
                                        if (!this.ScanJob.SkipUserInformation)
                                        {
                                            pageScanResult.ModifiedBy = page.LastModifiedBy();
                                        }

                                        // Page layout
                                        pageScanResult.PageLayout = page.PageLayout();
                                        pageScanResult.PageLayoutFile = page.PageLayoutFile().Replace(pageScanResult.SiteColUrl, "").Replace("/_catalogs/masterpage/", "");

                                        // Customization status                                        
                                        if (this.MasterPageGalleryCustomization == null)
                                        {
                                            this.MasterPageGalleryCustomization = new Dictionary<string, CustomizedPageStatus>();
                                        }

                                        // Load the file to check the customization status, only do this if the file was not loaded before for this site collection
                                        string layoutFile = page.PageLayoutFile();
                                        if (!string.IsNullOrEmpty(layoutFile))
                                        {
                                            Uri uri = new Uri(layoutFile);
                                            var url = page.PageLayoutFile().Replace($"{uri.Scheme}://{uri.DnsSafeHost}".ToLower(), "");
                                            if (!this.MasterPageGalleryCustomization.ContainsKey(url))
                                            {
                                                try
                                                {
                                                    var publishingPageLayout = cc.Site.RootWeb.GetFileByServerRelativeUrl(url);
                                                    cc.Load(publishingPageLayout);
                                                    cc.ExecuteQueryRetry();

                                                    this.MasterPageGalleryCustomization.Add(url, publishingPageLayout.CustomizedPageStatus);
                                                }
                                                catch (Exception ex)
                                                {
                                                    // eat potential exceptions
                                                }
                                            }

                                            // store the page layout customization status 
                                            if (this.MasterPageGalleryCustomization.TryGetValue(url, out CustomizedPageStatus pageStatus))
                                            {
                                                if (pageStatus == CustomizedPageStatus.Uncustomized)
                                                {
                                                    pageScanResult.PageLayoutWasCustomized = false;
                                                }
                                                else
                                                {
                                                    pageScanResult.PageLayoutWasCustomized = true;
                                                }

                                            }
                                            else
                                            {
                                                // If the file was not loaded for some reason then assume it was customized
                                                pageScanResult.PageLayoutWasCustomized = true;
                                            }
                                        }

                                        // Page audiences
                                        var audiences = page.Audiences();
                                        if (audiences != null)
                                        {
                                            pageScanResult.GlobalAudiences = audiences.GlobalAudiences;
                                            pageScanResult.SecurityGroupAudiences = audiences.SecurityGroups;
                                            pageScanResult.SharePointGroupAudiences = audiences.SharePointGroups;
                                        }

                                        // Contenttype
                                        pageScanResult.ContentType = page.ContentType.Name;
                                        pageScanResult.ContentTypeId = page.ContentType.Id.StringValue;

                                        // Get page web parts
                                        var pageAnalysis = page.WebParts(this.ScanJob.PageTransformation);
                                        if (pageAnalysis != null)
                                        {
                                            pageScanResult.WebParts = pageAnalysis.Item2;
                                        }

                                        // Persist publishing page scan results
                                        if (!this.ScanJob.PublishingPageScanResults.TryAdd(pageUrl, pageScanResult))
                                        {
                                            ScanError error = new ScanError()
                                            {
                                                Error = $"Could not add publishing page scan result for {pageScanResult.PageRelativeUrl}",
                                                SiteColUrl = this.SiteCollectionUrl,
                                                SiteURL = this.SiteUrl,
                                                Field1 = "PublishingAnalyzer",
                                                Field2 = pageScanResult.PageRelativeUrl,
                                            };
                                            this.ScanJob.ScanErrors.Push(error);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        ScanError error = new ScanError()
                                        {
                                            Error = ex.Message,
                                            SiteColUrl = this.SiteCollectionUrl,
                                            SiteURL = this.SiteUrl,
                                            Field1 = "MainPublishingPageAnalyzerLoop",
                                            Field2 = ex.StackTrace,
                                            Field3 = pageUrl
                                        };

                                        // Send error to telemetry to make scanner better
                                        if (this.ScanJob.ScannerTelemetry != null)
                                        {
                                            this.ScanJob.ScannerTelemetry.LogScanError(ex, error);
                                        }

                                        this.ScanJob.ScanErrors.Push(error);
                                        Console.WriteLine("Error for page {1}: {0}", ex.Message, pageUrl);
                                    }
                                }
                            }

                        }

                        // Persist publishing scan results
                        if (!this.ScanJob.PublishingWebScanResults.TryAdd(this.SiteUrl, scanResult))
                        {
                            ScanError error = new ScanError()
                            {
                                Error = $"Could not add publishing scan result for {this.SiteUrl}",
                                SiteColUrl = this.SiteCollectionUrl,
                                SiteURL = this.SiteUrl,
                                Field1 = "PublishingAnalyzer",
                            };
                            this.ScanJob.ScanErrors.Push(error);
                        }
                    }
                    catch(Exception ex)
                    {
                        ScanError error = new ScanError()
                        {
                            Error = ex.Message,
                            SiteColUrl = this.SiteCollectionUrl,
                            SiteURL = this.SiteUrl,
                            Field1 = "MainPublishingAnalyzerLoop",
                            Field2 = ex.StackTrace,
                        };

                        // Send error to telemetry to make scanner better
                        if (this.ScanJob.ScannerTelemetry != null)
                        {
                            this.ScanJob.ScannerTelemetry.LogScanError(ex, error);
                        }

                        this.ScanJob.ScanErrors.Push(error);
                        Console.WriteLine("Error for web {1}: {0}", ex.Message, this.SiteUrl);
                    }
                }
            }
            finally
            {
                this.StopTime = DateTime.Now;
            }

            // return the duration of this scan
            return new TimeSpan((this.StopTime.Subtract(this.StartTime).Ticks));
        }

        private static IEnumerable<VariationLabelEntity> GetVariationLabels(ClientContext context)
        {
            const string VARIATIONLABELSLISTID = "_VarLabelsListId";

            var variationLabels = new List<VariationLabelEntity>();
            // Get current web
            Web web = context.Web;
            web.EnsureProperty(w => w.ServerRelativeUrl);

            // Try to get _VarLabelsListId property from web property bag
            string variationLabelsListId = GetPropertyBagValue<string>(web, VARIATIONLABELSLISTID, string.Empty);

            if (!string.IsNullOrEmpty(variationLabelsListId))
            {
                var lists = context.Web.GetListsToScan(showHidden:true);
                Guid varRelationshipsListId = new Guid(variationLabelsListId);
                var variationLabelsList = lists.Where(p => p.Id.Equals(varRelationshipsListId)).FirstOrDefault();

                if (variationLabelsList != null)
                {
                    // Get the variationLabelsList list items
                    ListItemCollection collListItems = variationLabelsList.GetItems(CamlQuery.CreateAllItemsQuery());
                    context.Load(collListItems);
                    context.ExecuteQueryRetry();

                    foreach (var listItem in collListItems)
                    {
                        var label = new VariationLabelEntity();
                        label.Title = (string)listItem["Title"];
                        label.Description = (string)listItem["Description"];
                        label.FlagControlDisplayName = (string)listItem["Flag_x0020_Control_x0020_Display"];
                        label.Language = (string)listItem["Language"];
                        label.Locale = Convert.ToUInt32(listItem["Locale"]);
                        label.HierarchyCreationMode = (string)listItem["Hierarchy_x0020_Creation_x0020_M"];
                        label.IsSource = (bool)listItem["Is_x0020_Source"];
                        variationLabels.Add(label);
                    }
                }
            }
            return variationLabels;
        }

        private static T GetPropertyBagValue<T>(Web web, string key, T defaultValue)
        {
            web.EnsureProperty(p => p.AllProperties);
            
            if (web.AllProperties.FieldValues.ContainsKey(key))
            {
                return (T)web.AllProperties.FieldValues[key];
            }
            else
            {                
                return defaultValue;
            }
        }

        private string SanitizeXmlString(string xml)
        {
            // Turn into a list of bytes
            byte[] bytes = Encoding.UTF8.GetBytes(xml);
            List<byte> byteArray = bytes.ToList();

            // Check for preamble and delete it if needed
            foreach (byte singleByte in Encoding.UTF8.GetPreamble())
            {
                int pos = byteArray.IndexOf(singleByte);
                if (pos > -1)
                {
                    byteArray.RemoveAt(pos);
                }
            }

            // remove carriage returns and tabs
            xml = Encoding.UTF8.GetString(byteArray.ToArray());
            xml = xml.Replace("\\r", "");
            xml = xml.Replace("\\t", "");

            return xml;
        }

        private int ContinueScanning(ClientContext cc)
        {
            // Check site collection
            if (this.siteScanResult != null)
            {                
                Web web = cc.Web;

                // "Classic" publishing portal found
                if ((this.siteScanResult.WebTemplate == "BLANKINTERNET#0" || this.siteScanResult.WebTemplate == "ENTERWIKI#0" || 
                     this.siteScanResult.WebTemplate == "SRCHCEN#0" || this.siteScanResult.WebTemplate == "CMSPUBLISHING#0") &&
                    (this.siteScanResult.SitePublishingFeatureEnabled && this.siteScanResult.WebPublishingFeatureEnabled))
                {
                    var pagesLibrary = web.GetListsToScan().Where(p => p.BaseTemplate == 850).FirstOrDefault();
                    if (pagesLibrary != null)
                    {
                        // Take in account the "PageNotFoundError.aspx" default page. We want to assess if folks use publishing pages or not.
                        if (pagesLibrary.ItemCount > 1)
                        {
                            return pagesLibrary.ItemCount;
                        }
                    }

                    // always return a value in this case, if no pages found as this is a "classic" portal
                    return -1;
                }

                // Publishing enabled on non typical publishing portal site...check if there are pages in the Pages library
                if (this.siteScanResult.SitePublishingFeatureEnabled && this.siteScanResult.WebPublishingFeatureEnabled)
                {
                    
                    var pagesLibrary = web.GetListsToScan().Where(p => p.BaseTemplate == 850).FirstOrDefault();
                    if (pagesLibrary != null)
                    {
                        return pagesLibrary.ItemCount;
                    }
                }
            }
            return 0;
        }

        internal static Dictionary<string, PublishingSiteScanResult> GeneratePublishingSiteResults(Mode mode,
                                                                                                   ConcurrentDictionary<string, PublishingWebScanResult> webScanResults,
                                                                                                   ConcurrentDictionary<string, PublishingPageScanResult> pageScanResults)
        {
            Dictionary<string, PublishingSiteScanResult> siteScanResults = new Dictionary<string, PublishingSiteScanResult>(500);

            // bail out when no work todo
            if (!Options.IncludePublishing(mode) || webScanResults.Count == 0)
            {
                return null;
            }

            // iterate the web publishing results and consolidate into a single site level data line
            foreach (var item in webScanResults)
            {
                PublishingSiteScanResult siteResult = null;

                // Create or get the result instance
                if (!siteScanResults.ContainsKey(item.Value.SiteColUrl))
                {
                    siteResult = new PublishingSiteScanResult()
                    {
                        SiteColUrl = item.Value.SiteColUrl,
                        SiteURL = item.Value.SiteURL,
                        Classification = SiteComplexity.Simple
                    };
                    siteScanResults.Add(item.Value.SiteColUrl, siteResult);
                }
                else
                {
                    siteScanResults.TryGetValue(item.Value.SiteColUrl, out siteResult);
                }

                // Update the result instance
                siteResult.NumberOfWebs++;
                siteResult.NumberOfPages = siteResult.NumberOfPages + item.Value.PageCount;

                if (item.Value.SiteMasterPage != null && !siteResult.UsedSiteMasterPages.Contains(item.Value.SiteMasterPage))
                {
                    siteResult.UsedSiteMasterPages.Add(item.Value.SiteMasterPage);
                }

                if (item.Value.SystemMasterPage != null && !siteResult.UsedSystemMasterPages.Contains(item.Value.SystemMasterPage))
                {
                    siteResult.UsedSystemMasterPages.Add(item.Value.SystemMasterPage);
                }

                // If in a single site collection multiple languages are used then mark the publishing portal as complex
                if (!siteResult.UsedLanguages.Contains(item.Value.Language))
                {
                    siteResult.UsedLanguages.Add(item.Value.Language);
                }

                if (siteResult.UsedLanguages.Count > 1)
                {
                    siteResult.Classification = SiteComplexity.Complex;
                }
                
                // Check the classification based upon the web level data
                var webClassification = item.Value.WebClassification;
                if (webClassification > siteResult.Classification)
                {
                    siteResult.Classification = webClassification;
                }
            }

            // Iterate the publishing page results (if collected) and consolidate into a single site level data line
            if (Options.IncludePublishingWithPages(mode) && pageScanResults.Count > 0)
            {
                foreach (var item in pageScanResults)
                {
                    // Get the previously created record
                    siteScanResults.TryGetValue(item.Value.SiteColUrl, out PublishingSiteScanResult siteResult);

                    if (siteResult == null)
                    {
                        // Should not be possible...
                        continue;
                    }

                    // Update the result instance
                    if (item.Value.PageLayout != null && !siteResult.UsedPageLayouts.Contains(item.Value.PageLayout))
                    {
                        siteResult.UsedPageLayouts.Add(item.Value.PageLayout);
                    }

                    // Update last updated date               
                    if (item.Value.ModifiedAt != DateTime.MinValue && item.Value.ModifiedAt != DateTime.MaxValue)
                    {
                        if (siteResult.LastPageUpdateDate == null || item.Value.ModifiedAt > siteResult.LastPageUpdateDate)
                        {
                            siteResult.LastPageUpdateDate = item.Value.ModifiedAt;
                        }
                    }

                    // Update complexity based upon data for the page

                    // Customized page layouts
                    var pageClassification = SiteComplexity.Simple;
                    if (item.Value.PageLayoutWasCustomized)
                    {
                        pageClassification = SiteComplexity.Medium;
                    }
                    if (pageClassification > siteResult.Classification)
                    {
                        siteResult.Classification = pageClassification;
                    }

                    // Audiences used
                    pageClassification = SiteComplexity.Simple;
                    if ((item.Value.GlobalAudiences != null && item.Value.GlobalAudiences.Count > 0) || 
                        (item.Value.SecurityGroupAudiences != null && item.Value.SecurityGroupAudiences.Count > 0) || 
                        (item.Value.SharePointGroupAudiences != null && item.Value.SharePointGroupAudiences.Count > 0))
                    {
                        pageClassification = SiteComplexity.Medium;
                    }
                    if (pageClassification > siteResult.Classification)
                    {
                        siteResult.Classification = pageClassification;
                    }
                }
            }

            // Push back the site collection complexity level as a column of the web rows as that data is exported for the dashboard
            foreach (var item in webScanResults)
            {
                if (siteScanResults.TryGetValue(item.Value.SiteColUrl, out PublishingSiteScanResult site))
                {
                    item.Value.SiteClassification = site.Classification.ToString();
                }
            }

            return siteScanResults;
        }

    }
}
