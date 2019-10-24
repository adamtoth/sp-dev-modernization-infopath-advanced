﻿using Microsoft.SharePoint.Client;
using SharePoint.Modernization.Scanner.Results;
using SharePoint.Scanning.Framework;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePoint.Modernization.Scanner.Analyzers
{
    public class ListAnalyzer : BaseAnalyzer
    {

        #region Construction
        /// <summary>
        /// List analyzer construction
        /// </summary>
        /// <param name="url">Url of the web to be analyzed</param>
        /// <param name="siteColUrl">Url of the site collection hosting this web</param>
        public ListAnalyzer(string url, string siteColUrl, ModernizationScanJob scanJob) : base(url, siteColUrl, scanJob)
        {
        }
        #endregion

        /// <summary>
        /// Analyze the web
        /// </summary>
        /// <param name="cc">ClientContext of the web to be analyzed</param>
        /// <returns>Duration of the analysis</returns>
        public override TimeSpan Analyze(ClientContext cc)
        {
            try
            {
                base.Analyze(cc);

                var baseUri = new Uri(this.SiteUrl);
                var webAppUrl = baseUri.Scheme + "://" + baseUri.Host;

                var lists = cc.Web.GetListsToScan();

                foreach (var list in lists)
                {
                    try
                    {
                        this.ScanJob.IncreaseScannedLists();

                        ListScanResult listScanData;
                        if (list.DefaultViewUrl.ToLower().Contains(".aspx"))
                        {
                            File file = cc.Web.GetFileByServerRelativeUrl(list.DefaultViewUrl);
                            listScanData = file.ModernCompatability(list, ref this.ScanJob.ScanErrors);
                        }
                        else
                        {
                            listScanData = new ListScanResult()
                            {
                                BlockedByNotBeingAbleToLoadPage = true
                            };
                        }

                        if (listScanData != null && !listScanData.WorksInModern)
                        {
                            if (this.ScanJob.ExcludeListsOnlyBlockedByOobReasons && listScanData.OnlyBlockedByOOBReasons)
                            {
                                continue;
                            }

                            listScanData.SiteURL = this.SiteUrl;
                            listScanData.ListUrl = $"{webAppUrl}{list.DefaultViewUrl}";
                            listScanData.SiteColUrl = this.SiteCollectionUrl;
                            listScanData.ListTitle = list.Title;

                            if (!this.ScanJob.ListScanResults.TryAdd($"{Guid.NewGuid().ToString()}{webAppUrl}{list.DefaultViewUrl}", listScanData))
                            {
                                ScanError error = new ScanError()
                                {
                                    Error = $"Could not add list scan result for {webAppUrl}{list.DefaultViewUrl} from web scan of {this.SiteUrl}",
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    Field1 = "ListAnalyzer",
                                    Field2 = $"{webAppUrl}{list.DefaultViewUrl}"
                                };
                                this.ScanJob.ScanErrors.Push(error);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ScanError error = new ScanError()
                        {
                            Error = ex.Message,
                            SiteColUrl = this.SiteCollectionUrl,
                            SiteURL = this.SiteUrl,
                            Field1 = "MainListAnalyzerLoop",
                            Field2 = ex.StackTrace,
                            Field3 = $"{webAppUrl}{list.DefaultViewUrl}"
                        };

                        // Send error to telemetry to make scanner better
                        if (this.ScanJob.ScannerTelemetry != null)
                        {
                            this.ScanJob.ScannerTelemetry.LogScanError(ex, error);
                        }

                        this.ScanJob.ScanErrors.Push(error);
                        Console.WriteLine("Error for page {1}: {0}", ex.Message, $"{webAppUrl}{list.DefaultViewUrl}");
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

        internal static List<string> GenerateSitesWithCodeCustomizationsResults(ConcurrentDictionary<string, ListScanResult> listScanResults)
        {
            List<string> sitesWithCodeCustomizationsResults = new List<string>(500);

            foreach(var list in listScanResults)
            {
                if (list.Value.BlockedAtSiteLevel ||
                    list.Value.BlockedAtWebLevel || 
                    list.Value.XsltViewWebPartCompatibility.BlockedByJSLink ||
                    list.Value.XsltViewWebPartCompatibility.BlockedByJSLinkField ||
                    list.Value.XsltViewWebPartCompatibility.BlockedByListCustomAction ||
                    list.Value.XsltViewWebPartCompatibility.BlockedByXsl ||
                    list.Value.XsltViewWebPartCompatibility.BlockedByXslLink)
                {
                    if (!sitesWithCodeCustomizationsResults.Contains(list.Value.SiteColUrl))
                    {
                        sitesWithCodeCustomizationsResults.Add(list.Value.SiteColUrl);
                    }
                }
            }

            return sitesWithCodeCustomizationsResults;
        }
    }
}
