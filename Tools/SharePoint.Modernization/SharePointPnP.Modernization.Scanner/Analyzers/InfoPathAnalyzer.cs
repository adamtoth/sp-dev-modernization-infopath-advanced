using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.VisualBasic.FileIO;
using SharePoint.Modernization.Scanner.Results;
using SharePoint.Scanning.Framework;

namespace SharePoint.Modernization.Scanner.Analyzers
{
    public class InfoPathAnalyzer: BaseAnalyzer
    {

        private static readonly string FormBaseContentType = "0x010101";

        #region Construction
        /// <summary>
        /// InfoPath analyzer construction
        /// </summary>
        /// <param name="url">Url of the web to be analyzed</param>
        /// <param name="siteColUrl">Url of the site collection hosting this web</param>
        /// <param name="scanJob">Job that launched this analyzer</param>
        public InfoPathAnalyzer(string url, string siteColUrl, ModernizationScanJob scanJob) : base(url, siteColUrl, scanJob)
        {            
        }
        #endregion

        #region Analysis
        /// <summary>
        /// Analyses a web for it's workflow usage
        /// </summary>
        /// <param name="cc">ClientContext instance used to retrieve workflow data</param>
        /// <returns>Duration of the workflow analysis</returns>
        public override TimeSpan Analyze(ClientContext cc)
        {
            try
            {
                base.Analyze(cc);

                var baseUri = new Uri(this.SiteUrl);
                var webAppUrl = baseUri.Scheme + "://" + baseUri.Host;

                var lists = cc.Web.GetListsToScan(showHidden: true);

                foreach (var list in lists)
                {
                    // Skip system lists, except for Converted Forms and Form Templates libraries
                    if ((list.IsSystemList || list.IsCatalog || list.IsSiteAssetsLibrary || list.IsEnterpriseGalleryLibrary) &&
                        ((int)list.BaseTemplate != 10102 || !string.Equals(list.RootFolder.Name, "FormServerTemplates", StringComparison.InvariantCultureIgnoreCase)))
                    {
                        continue;
                    }

                    if ((int)list.BaseTemplate == 10102)
                    {
                        // Converted Forms library
                        // Stores converted InfoPath forms for browser rendering. This can become bloated 
                        // with many versions stored after every republish. Recommended to clean this out 
                        // after migrating away from InfoPath                       
                        if (list.ItemCount > 0)
                        {
                            InfoPathScanResult infoPathScanResult = new InfoPathScanResult()
                            {
                                SiteColUrl = this.SiteCollectionUrl,
                                SiteURL = this.SiteUrl,
                                InfoPathUsage = "ConvertedFormLibrary",
                                ListTitle = list.Title,
                                ListId = list.Id,
                                ListUrl = list.RootFolder.ServerRelativeUrl,
                                Enabled = true,
                                InfoPathTemplate = string.Empty,
                                InfoPathTemplateUrl = string.Empty,
                                ItemCount = list.ItemCount,
                                LastItemUserModifiedDate = list.LastItemUserModifiedDate,
                            };
                            if (!this.ScanJob.InfoPathScanResults.TryAdd($"{infoPathScanResult.SiteURL}.{Guid.NewGuid()}", infoPathScanResult))
                            {
                                ScanError error = new ScanError()
                                {
                                    Error = $"Could not add formlibrary InfoPath scan result for {infoPathScanResult.SiteColUrl} and list {infoPathScanResult.ListUrl}",
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    Field1 = "InfoPathAnalyzer",
                                };
                                this.ScanJob.ScanErrors.Push(error);
                            }
                        }
                        continue;

                    }

                    if (string.Equals(list.RootFolder.Name, "FormServerTemplates", StringComparison.InvariantCultureIgnoreCase))
                    {
                        // Form Templates library
                        // Any templates in here indicate sandbox solution/admin approved forms
                        // Possibly from an on-prem migration?
                        if (list.ItemCount > 0)
                        {
                            InfoPathScanResult infoPathScanResult = new InfoPathScanResult()
                            {
                                SiteColUrl = this.SiteCollectionUrl,
                                SiteURL = this.SiteUrl,
                                InfoPathUsage = "FormServerTemplates",
                                ListTitle = list.Title,
                                ListId = list.Id,
                                ListUrl = list.RootFolder.ServerRelativeUrl,
                                Enabled = true,
                                InfoPathTemplate = string.Empty,
                                InfoPathTemplateUrl = string.Empty,
                                ItemCount = list.ItemCount,
                                LastItemUserModifiedDate = list.LastItemUserModifiedDate,
                            };
                            if (!this.ScanJob.InfoPathScanResults.TryAdd($"{infoPathScanResult.SiteURL}.{Guid.NewGuid()}", infoPathScanResult))
                            {
                                ScanError error = new ScanError()
                                {
                                    Error = $"Could not add formlibrary InfoPath scan result for {infoPathScanResult.SiteColUrl} and list {infoPathScanResult.ListUrl}",
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    Field1 = "InfoPathAnalyzer",
                                };
                                this.ScanJob.ScanErrors.Push(error);
                            }
                        }
                        continue;
                    }


                    if (list.BaseType == BaseType.DocumentLibrary)
                    {

                        if (!String.IsNullOrEmpty(list.DocumentTemplateUrl) && list.DocumentTemplateUrl.EndsWith(".xsn", StringComparison.InvariantCultureIgnoreCase))
                        {
                            // Library using an InfoPath file as New Item template
                            InfoPathScanResult infoPathScanResult = new InfoPathScanResult()
                            {
                                SiteColUrl = this.SiteCollectionUrl,
                                SiteURL = this.SiteUrl,
                                InfoPathUsage = "FormLibrary",
                                ListTitle = list.Title,
                                ListId = list.Id,
                                ListUrl = list.RootFolder.ServerRelativeUrl,
                                Enabled = true,
                                InfoPathTemplate = list.DocumentTemplateUrl,
                                InfoPathTemplateUrl = this.SiteUrl.TrimEnd('/') + "/_layouts/16/download.aspx?SourceUrl=" + list.DocumentTemplateUrl,
                                ItemCount = list.ItemCount,
                                LastItemUserModifiedDate = list.LastItemUserModifiedDate,
                            };
                            // Root relative URL for the template
                            var templateFile = cc.Web.GetFileByServerRelativeUrl(list.DocumentTemplateUrl);
                            infoPathScanResult = ScrapeXsn(cc, infoPathScanResult, templateFile);

                            if (!this.ScanJob.InfoPathScanResults.TryAdd($"{infoPathScanResult.SiteURL}.{Guid.NewGuid()}", infoPathScanResult))
                            {
                                ScanError error = new ScanError()
                                {
                                    Error = $"Could not add formlibrary InfoPath scan result for {infoPathScanResult.SiteColUrl} and list {infoPathScanResult.ListUrl}",
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    Field1 = "InfoPathAnalyzer",
                                };
                                this.ScanJob.ScanErrors.Push(error);
                            }

                        }

                        // Iterate through all content types - we may not have a template assigned to the library itself,
                        // only to content types
                        cc.Load(list, p => p.ContentTypes.Include(i => i.Name, i => i.DocumentTemplate, i => i.DocumentTemplateUrl, i => i.Id));
                        cc.ExecuteQueryRetry();

                        foreach (var contentType in list.ContentTypes)
                        {
                            if ((!string.IsNullOrEmpty(contentType.DocumentTemplateUrl) && 
                                contentType.DocumentTemplateUrl.EndsWith(".xsn", StringComparison.InvariantCultureIgnoreCase)) || 
                                contentType.Id.StringValue.StartsWith(FormBaseContentType))
                            {
                                InfoPathScanResult ctInfoPathScanResult = new InfoPathScanResult()
                                {
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    InfoPathUsage = "FormLibraryContentType",
                                    ListTitle = list.Title,
                                    ListId = list.Id,
                                    ListUrl = list.RootFolder.ServerRelativeUrl,
                                    Enabled = true,
                                    InfoPathTemplate = contentType.DocumentTemplateUrl,
                                    InfoPathTemplateUrl = this.SiteUrl.TrimEnd('/') + "/_layouts/16/download.aspx?SourceUrl=" + contentType.DocumentTemplateUrl,
                                    ItemCount = list.ItemCount,
                                    ContentTypeName = contentType.Name,
                                    LastItemUserModifiedDate = list.LastItemUserModifiedDate,
                                };
                                // Root relative URL for the template
                                var templateFile = cc.Web.GetFileByServerRelativeUrl(contentType.DocumentTemplateUrl);
                                ctInfoPathScanResult = ScrapeXsn(cc, ctInfoPathScanResult, templateFile);
                                if (!this.ScanJob.InfoPathScanResults.TryAdd($"{ctInfoPathScanResult.SiteURL}.{Guid.NewGuid()}", ctInfoPathScanResult))
                                {
                                    ScanError error = new ScanError()
                                    {
                                        Error = $"Could not add formlibrary InfoPath scan result for {ctInfoPathScanResult.SiteColUrl} and list {ctInfoPathScanResult.ListUrl}",
                                        SiteColUrl = this.SiteCollectionUrl,
                                        SiteURL = this.SiteUrl,
                                        Field1 = "InfoPathAnalyzer",
                                    };
                                    this.ScanJob.ScanErrors.Push(error);
                                }
                            }
                        }


                    }
                    else if (list.BaseType == BaseType.GenericList)
                    {
                        try
                        {
                            Folder folder = cc.Web.GetFolderByServerRelativeUrl($"{list.RootFolder.ServerRelativeUrl}/Item");
                            cc.Load(folder, p => p.Properties);
                            cc.ExecuteQueryRetry();

                            if (folder.Properties.FieldValues.ContainsKey("_ipfs_infopathenabled") && folder.Properties.FieldValues.ContainsKey("_ipfs_solutionName"))
                            {
                                bool infoPathEnabled = true;
                                if (bool.TryParse(folder.Properties.FieldValues["_ipfs_infopathenabled"].ToString(), out bool infoPathEnabledParsed))
                                {
                                    infoPathEnabled = infoPathEnabledParsed;
                                }

                                // List with an InfoPath customization
                                InfoPathScanResult infoPathScanResult = new InfoPathScanResult()
                                {
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    InfoPathUsage = "CustomForm",
                                    ListTitle = list.Title,
                                    ListId = list.Id,
                                    ListUrl = list.RootFolder.ServerRelativeUrl,
                                    Enabled = infoPathEnabled,
                                    InfoPathTemplate = folder.Properties.FieldValues["_ipfs_solutionName"].ToString(),
                                    InfoPathTemplateUrl = this.SiteUrl.TrimEnd('/') + "/_layouts/16/download.aspx?SourceUrl=" + cc.Web.Url.TrimEnd('/') + "/Lists/" + list.RootFolder.Name + "/" + folder.Properties.FieldValues["_ipfs_solutionName"].ToString(),
                                    ItemCount = list.ItemCount,
                                    LastItemUserModifiedDate = list.LastItemUserModifiedDate,
                                };

                                // Run the scraper
                                var templateFile = list.RootFolder.GetFile($"item/{infoPathScanResult.InfoPathTemplate}");
                                infoPathScanResult = ScrapeXsn(cc, infoPathScanResult, templateFile);

                                if (!this.ScanJob.InfoPathScanResults.TryAdd($"{infoPathScanResult.SiteURL}.{Guid.NewGuid()}", infoPathScanResult))
                                {
                                    ScanError error = new ScanError()
                                    {
                                        Error = $"Could not add customform InfoPath scan result for {infoPathScanResult.SiteColUrl} and list {infoPathScanResult.ListUrl}",
                                        SiteColUrl = this.SiteCollectionUrl,
                                        SiteURL = this.SiteUrl,
                                        Field1 = "InfoPathAnalyzer",
                                    };
                                    this.ScanJob.ScanErrors.Push(error);
                                }


                            }
                        }
                        catch (ServerException ex)
                        {
                            if (((ServerException)ex).ServerErrorTypeName == "System.IO.FileNotFoundException")
                            {
                                // Ignore
                            }
                            else
                            {
                                ScanError error = new ScanError()
                                {
                                    Error = ex.Message,
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    Field1 = "InfoPathAnalyzer",
                                    Field2 = ex.StackTrace,
                                    Field3 = $"{webAppUrl}{list.DefaultViewUrl}"
                                };

                                // Send error to telemetry to make scanner better
                                if (this.ScanJob.ScannerTelemetry != null)
                                {
                                    this.ScanJob.ScannerTelemetry.LogScanError(ex, error);
                                }

                                this.ScanJob.ScanErrors.Push(error);
                                Console.WriteLine("Error during InfoPath analysis for list {1}: {0}", ex.Message, $"{webAppUrl}{list.DefaultViewUrl}");
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                this.StopTime = DateTime.Now;
            }

            // return the duration of this scan
            return new TimeSpan((this.StopTime.Subtract(this.StartTime).Ticks));
        }

        public InfoPathScanResult ScrapeXsn(ClientContext cc, InfoPathScanResult infoPathScanResult, Microsoft.SharePoint.Client.File templateFile)
        {
            // Run the scraper
            System.IO.DirectoryInfo xsnDir = System.IO.Directory.CreateDirectory(this.ScanJob.OutputFolder + "\\InfoPathXsnFiles");
            var now = DateTime.Now;
            string timeStampAsString = now.ToFileTime().ToString();
            string tempFilePath = xsnDir.FullName + "\\" + timeStampAsString + ".xsn";

            ClientResult<System.IO.Stream> xsnFile = templateFile.OpenBinaryStream();
            cc.Load(templateFile);
            cc.ExecuteQueryRetry();

            using (var fs = System.IO.File.Create(tempFilePath))
            {
                xsnFile.Value.CopyTo(fs);
            }

            infoPathScanResult.DownloadedXsnId = timeStampAsString;

            using (Process scraperProcess = new Process())
            {
                scraperProcess.StartInfo.UseShellExecute = false;
                scraperProcess.StartInfo.FileName = "InfoPathScraper.exe";
                scraperProcess.StartInfo.WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory;
                scraperProcess.StartInfo.Arguments = "/csv /file \"" + tempFilePath + "\" /outfile \"" + tempFilePath.Replace(".xsn",".csv") + "\"";
                scraperProcess.StartInfo.CreateNoWindow = false;
                scraperProcess.Start();
                scraperProcess.WaitForExit();
            }

            // Scraper outputs CSV file
            // Each row is a different property, with variable number of columns
            // You have to parse each line independently, as each line has different schema
            // There are no column headers
            // The TextFieldParser from Microsoft.VisualBasic.dll does a good job of basic row and field enumeration
            using (TextFieldParser parser = new TextFieldParser(tempFilePath.Replace(".xsn", ".csv")))
            {
                parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
                parser.SetDelimiters(",");

                bool modeFound = false;
                bool productVersionFound = false;
                bool controlFound = false;
                bool soapConnectionFound = false;
                bool dbConnectionFound = false;
                bool restConnectionFound = false;

                while (!parser.EndOfData)
                {

                    //Process row
                    string[] fields = parser.ReadFields();

                    // store a boolean flags for each column you are looking for
                    // if found, the values will be in subsequent columns
                    // Ater parsing subsequent columns, set the flag to true to avoid looking for it
                    foreach (string field in fields)
                    {
                        if (modeFound)
                        {
                            infoPathScanResult.Mode = field;
                            modeFound = false;
                            break;
                        }
                        else
                        {
                            if (string.Equals(field, "Mode"))
                            {
                                modeFound = true;
                            }
                        }
                        if (productVersionFound)
                        {
                            infoPathScanResult.ProductVersion = field;
                            productVersionFound = false;
                            break;
                        }
                        else
                        {
                            if (string.Equals(field, "ProductVersion"))
                            {
                                productVersionFound = true;
                            }
                        }
                        if (controlFound)
                        {
                            if (field == "{{61e40d31-993d-4777-8fa0-19ca59b6d0bb}}")
                            {
                                infoPathScanResult.HasPersonField = true;
                            }
                            else if (string.Equals(field, "entitypicker", StringComparison.InvariantCultureIgnoreCase))
                            {
                                infoPathScanResult.HasExternalField = true;
                            }
                            else if (string.Equals(field, "RepeatingTable", StringComparison.InvariantCultureIgnoreCase))
                            {
                                infoPathScanResult.HasRepeatingTable = true;
                            }
                            else if (string.Equals(field, "RepeatingSection", StringComparison.InvariantCultureIgnoreCase))
                            {
                                infoPathScanResult.HasRepeatingSection = true;
                            }
                            else if (string.Equals(field, "choicegroup", StringComparison.InvariantCultureIgnoreCase))
                            {
                                infoPathScanResult.HasChoiceGroup = true;
                            }
                            else if (string.Equals(field, "choicegrouprepeating", StringComparison.InvariantCultureIgnoreCase))
                            {
                                infoPathScanResult.HasRepeatingChoiceGroup = true;
                            }
                            else if (string.Equals(field, "choiceterm", StringComparison.InvariantCultureIgnoreCase))
                            {
                                infoPathScanResult.HasChoiceSection = true;
                            }
                            else if (string.Equals(field, "choicetermrepeating", StringComparison.InvariantCultureIgnoreCase))
                            {
                                infoPathScanResult.HasRepeatingChoiceGroup = true;
                            }
                            else if (string.Equals(field, "SignatureLine", StringComparison.InvariantCultureIgnoreCase))
                            {
                                infoPathScanResult.HasDigitalSignature = true;
                            }
                            else if (string.Equals(field, "inkpicture", StringComparison.InvariantCultureIgnoreCase))
                            {
                                infoPathScanResult.HasInk = true;
                            }
                            else if (string.Equals(field, "SignatureLine", StringComparison.InvariantCultureIgnoreCase))
                            {
                                infoPathScanResult.HasDigitalSignature = true;
                            }
                            else if (string.Equals(field, "PageBreak", StringComparison.InvariantCultureIgnoreCase))
                            {
                                infoPathScanResult.HasPageBreak = true;
                            }
                            controlFound = false;
                            break;
                        }
                        else
                        {
                            if (string.Equals(field, "Control"))
                            {
                                controlFound = true;
                            }
                        }
                        if (string.Equals(field, "SoapConnection"))
                        {
                            infoPathScanResult.HasSOAPConnection = true;
                            break;
                        }
                        if (string.Equals(field, "DBConnection"))
                        {
                            infoPathScanResult.HasDBConnection = true;
                            break;
                        }
                        if (string.Equals(field, "RESTConnection"))
                        {
                            infoPathScanResult.HasRESTConnection = true;
                            break;
                        }
                    }
                }
            }


            return infoPathScanResult;
        }

        #endregion

    }
}
