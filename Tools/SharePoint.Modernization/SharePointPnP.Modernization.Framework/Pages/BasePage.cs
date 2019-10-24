﻿using AngleSharp.Dom;
using AngleSharp.Dom.Html;
using AngleSharp.Parser.Html;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using SharePointPnP.Modernization.Framework.Entities;
using SharePointPnP.Modernization.Framework.Extensions;
using SharePointPnP.Modernization.Framework.Telemetry;
using SharePointPnP.Modernization.Framework.Transform;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace SharePointPnP.Modernization.Framework.Pages
{

    /// <summary>
    /// Base class for the page analyzers
    /// </summary>
    public abstract class BasePage : BaseTransform
    {
        internal class WebPartPlaceHolder
        {
            public string Id { get; set; }
            public string ControlId { get; set; }
            public int Row { get; set; }
            public int Column { get; set; }
            public int Order { get; set; }
            public WebPartDefinition WebPartDefinition { get; set; }
            public string WebPartXmlOnPremises { get; set; }
            public ClientResult<string> WebPartXml { get; set; }

            public string WebPartType { get; set; }
        }

        internal HtmlParser parser;
        private const string webPartMarkerString = "[[WebPartMarker]]";

        public ListItem page;
        public File pageFile;
        public ClientContext cc;
        public PageTransformation pageTransformation;

        #region construction
        /// <summary>
        /// Constructs the base page class instance
        /// </summary>
        /// <param name="page">page ListItem</param>
        /// <param name="pageFile">page File</param>
        /// <param name="pageTransformation">page transformation model to use for extraction or transformation</param>
        public BasePage(ListItem page, File pageFile, PageTransformation pageTransformation, IList<ILogObserver> logObservers = null)
        {
            // Register observers
            if (logObservers != null)
            {
                foreach (var observer in logObservers)
                {
                    base.RegisterObserver(observer);
                }
            }

            this.page = page;
            this.pageFile = pageFile;

            if (page != null)
            {
                this.cc = (page.Context as ClientContext);
            }
            else if (pageFile != null)
            {
                this.cc = (pageFile.Context as ClientContext);
            }

            this.cc.RequestTimeout = Timeout.Infinite;

            this.pageTransformation = pageTransformation;
            this.parser = new HtmlParser();
        }
        #endregion

        /// <summary>
        /// Get's the type of the web part
        /// </summary>
        /// <param name="webPartXml">Web part xml to analyze</param>
        /// <returns>Type of the web part as fully qualified name</returns>
        public string GetType(string webPartXml)
        {
            string type = "Unknown";

            if (!string.IsNullOrEmpty(webPartXml))
            {
                var xml = XElement.Parse(webPartXml);
                var xmlns = xml.XPathSelectElement("*").GetDefaultNamespace();
                if (xmlns.NamespaceName.Equals("http://schemas.microsoft.com/WebPart/v3", StringComparison.InvariantCultureIgnoreCase))
                {
                    type = xml.Descendants(xmlns + "type").FirstOrDefault().Attribute("name").Value;
                }
                else if (xmlns.NamespaceName.Equals("http://schemas.microsoft.com/WebPart/v2", StringComparison.InvariantCultureIgnoreCase))
                {
                    type = $"{xml.Descendants(xmlns + "TypeName").FirstOrDefault().Value}, {xml.Descendants(xmlns + "Assembly").FirstOrDefault().Value}";
                }
            }

            return type;
        }

        internal void AnalyzeWikiContentBlock(List<WebPartEntity> webparts, IHtmlDocument htmlDoc, List<WebPartPlaceHolder> webPartsToRetrieve, int rowCount, int colCount, int startOrder, IElement content)
        {
            // Drop elements which we anyhow can't transform and/or which are stripped out from RTE
            CleanHtml(content, htmlDoc);

            StringBuilder textContent = new StringBuilder();
            int order = startOrder;
            foreach (var node in content.ChildNodes)
            {
                // Do we find a web part inside...
                if (((node as IHtmlElement) != null) && ContainsWebPart(node as IHtmlElement))
                {
                    var extraText = StripWebPart(node as IHtmlElement);
                    string extraTextAfterWebPart = null;
                    string extraTextBeforeWebPart = null;
                    if (!string.IsNullOrEmpty(extraText))
                    {
                        // Should be, but checking anyhow
                        int webPartMarker = extraText.IndexOf(webPartMarkerString);
                        if (webPartMarker > -1)
                        {
                            extraTextBeforeWebPart = extraText.Substring(0, webPartMarker);
                            extraTextAfterWebPart = extraText.Substring(webPartMarker + webPartMarkerString.Length);

                            // there could have been multiple web parts in a row (we don't support text inbetween them for now)...strip the remaining markers
                            extraTextBeforeWebPart = extraTextBeforeWebPart.Replace(webPartMarkerString, "");
                            extraTextAfterWebPart = extraTextAfterWebPart.Replace(webPartMarkerString, "");
                        }
                    }

                    if (!string.IsNullOrEmpty(extraTextBeforeWebPart))
                    {
                        textContent.AppendLine(extraTextBeforeWebPart);
                    }

                    // first insert text part (if it was available)
                    if (!string.IsNullOrEmpty(textContent.ToString()))
                    {
                        order++;
                        webparts.Add(CreateWikiTextPart(textContent.ToString(), rowCount, colCount, order));
                        textContent.Clear();
                    }

                    // then process the web part
                    order++;
                    Regex regexClientIds = new Regex(@"id=\""div_(?<ControlId>(\w|\-)+)");
                    if (regexClientIds.IsMatch((node as IHtmlElement).OuterHtml))
                    {
                        foreach (Match webPartMatch in regexClientIds.Matches((node as IHtmlElement).OuterHtml))
                        {
                            // Store the web part we need, will be retrieved afterwards to optimize performance
                            string serverSideControlId = webPartMatch.Groups["ControlId"].Value;
                            var serverSideControlIdToSearchFor = $"g_{serverSideControlId.Replace("-", "_")}";
                            webPartsToRetrieve.Add(new WebPartPlaceHolder() { ControlId = serverSideControlIdToSearchFor, Id = serverSideControlId, Row = rowCount, Column = colCount, Order = order });
                        }
                    }

                    // Process the extra text that was positioned after the web part (if any)
                    if (!string.IsNullOrEmpty(extraTextAfterWebPart))
                    {
                        textContent.AppendLine(extraTextAfterWebPart);
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(node.TextContent.Trim()) && node.TextContent.Trim() == "\n")
                    {
                        // ignore, this one is typically added after a web part
                    }
                    else
                    {
                        if (node.HasChildNodes)
                        {
                            textContent.AppendLine((node as IHtmlElement).OuterHtml);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(node.TextContent.Trim()))
                            {
                                textContent.AppendLine(node.TextContent);
                            }
                            else
                            {
                                if (node.NodeName.Equals("br", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    textContent.AppendLine("<BR>");
                                }
                                // given that wiki html can contain embedded images and videos while not having child nodes we need include these.
                                // case: img/iframe tag as "only" element to evaluate (e.g. first element in the contenthost)
                                else if (node.NodeName.Equals("img", StringComparison.InvariantCultureIgnoreCase) ||
                                         node.NodeName.Equals("iframe", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    textContent.AppendLine((node as IHtmlElement).OuterHtml);
                                }
                            }
                        }
                    }
                }
            }

            // there was only one text part
            if (!string.IsNullOrEmpty(textContent.ToString()))
            {
                // insert text part to the web part collection
                order++;
                webparts.Add(CreateWikiTextPart(textContent.ToString(), rowCount, colCount, order));
            }
        }

        /// <summary>
        /// Load Web Parts from Wiki Content page on Online Server
        /// </summary>
        /// <param name="webparts"></param>
        /// <param name="wikiPage"></param>
        /// <param name="webPartsToRetrieve"></param>
        internal void LoadWebPartsInWikiContentFromServer(List<WebPartEntity> webparts, File wikiPage, List<WebPartPlaceHolder> webPartsToRetrieve)
        {
            // Load web part manager and use it to load each web part
            LimitedWebPartManager limitedWPManager = wikiPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            cc.Load(limitedWPManager);

            foreach (var webPartToRetrieve in webPartsToRetrieve)
            {
                // Check if the web part was loaded when we loaded the web parts collection via the web part manager
                if (!Guid.TryParse(webPartToRetrieve.Id, out Guid webPartToRetrieveGuid))
                {
                    // Skip since guid is not valid
                    continue;
                }

                // Sometimes the returned wiki html contains web parts which are not anymore on the page...using the ExceptionHandlingScope 
                // we can handle these errors server side while just doing a single roundtrip
                var scope = new ExceptionHandlingScope(cc);
                using (scope.StartScope())
                {
                    using (scope.StartTry())
                    {
                        webPartToRetrieve.WebPartDefinition = limitedWPManager.WebParts.GetByControlId(webPartToRetrieve.ControlId);
                        cc.Load(webPartToRetrieve.WebPartDefinition, wp => wp.Id, wp => wp.WebPart.ExportMode, wp => wp.WebPart.Title, wp => wp.WebPart.ZoneIndex, wp => wp.WebPart.IsClosed, wp => wp.WebPart.Hidden, wp => wp.WebPart.Properties);
                    }
                    using (scope.StartCatch())
                    {

                    }
                }
            }
            cc.ExecuteQueryRetry();


            // Load the web part XML for the web parts that do allow it
            bool isDirty = false;
            foreach (var webPartToRetrieve in webPartsToRetrieve)
            {
                // Important to only process the web parts that did not return an error in the previous server call
                if (webPartToRetrieve.WebPartDefinition != null && (webPartToRetrieve.WebPartDefinition.ServerObjectIsNull.HasValue && webPartToRetrieve.WebPartDefinition.ServerObjectIsNull.Value == false))
                {
                    // Retry to load the properties, sometimes they're not retrieved
                    webPartToRetrieve.WebPartDefinition.EnsureProperty(wp => wp.Id);
                    webPartToRetrieve.WebPartDefinition.WebPart.EnsureProperties(wp => wp.ExportMode, wp => wp.Title, wp => wp.ZoneIndex, wp => wp.IsClosed, wp => wp.Hidden, wp => wp.Properties);

                    if (webPartToRetrieve.WebPartDefinition.WebPart.ExportMode == WebPartExportMode.All)
                    {
                        webPartToRetrieve.WebPartXml = limitedWPManager.ExportWebPart(webPartToRetrieve.WebPartDefinition.Id);
                        isDirty = true;
                    }
                }
            }
            if (isDirty)
            {
                cc.ExecuteQueryRetry();
            }

            // Determine the web part type and store it in the web parts array
            foreach (var webPartToRetrieve in webPartsToRetrieve)
            {
                if (webPartToRetrieve.WebPartDefinition != null && (webPartToRetrieve.WebPartDefinition.ServerObjectIsNull.HasValue && webPartToRetrieve.WebPartDefinition.ServerObjectIsNull.Value == false))
                {
                    // Important to only process the web parts that did not return an error in the previous server call
                    if (webPartToRetrieve.WebPartDefinition.WebPart.ExportMode != WebPartExportMode.All)
                    {
                        // Use different approach to determine type as we can't export the web part XML without indroducing a change
                        webPartToRetrieve.WebPartType = GetTypeFromProperties(webPartToRetrieve.WebPartDefinition.WebPart.Properties.FieldValues);
                    }
                    else
                    {
                        webPartToRetrieve.WebPartType = GetType(webPartToRetrieve.WebPartXml.Value);
                    }

                    webparts.Add(new WebPartEntity()
                    {
                        Title = webPartToRetrieve.WebPartDefinition.WebPart.Title,
                        Type = webPartToRetrieve.WebPartType,
                        Id = webPartToRetrieve.WebPartDefinition.Id,
                        ServerControlId = webPartToRetrieve.Id,
                        Row = webPartToRetrieve.Row,
                        Column = webPartToRetrieve.Column,
                        Order = webPartToRetrieve.Order,
                        ZoneId = "",
                        ZoneIndex = (uint)webPartToRetrieve.WebPartDefinition.WebPart.ZoneIndex,
                        IsClosed = webPartToRetrieve.WebPartDefinition.WebPart.IsClosed,
                        Hidden = webPartToRetrieve.WebPartDefinition.WebPart.Hidden,
                        Properties = Properties(webPartToRetrieve.WebPartDefinition.WebPart.Properties.FieldValues, webPartToRetrieve.WebPartType, webPartToRetrieve.WebPartXml == null ? "" : webPartToRetrieve.WebPartXml.Value),
                    });
                }
            }
        }

        /// <summary>
        /// Load Web Parts from Wiki Content page on On-Premises Server
        /// </summary>
        /// <param name="webparts"></param>
        /// <param name="wikiPage"></param>
        /// <param name="webPartsToRetrieve"></param>
        internal void LoadWebPartsInWikiContentFromOnPremisesServer(List<WebPartEntity> webparts, File wikiPage, List<WebPartPlaceHolder> webPartsToRetrieve)
        {
            // Load web part manager and use it to load each web part
            LimitedWebPartManager limitedWPManager = wikiPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            cc.Load(limitedWPManager);

            IEnumerable<WebPartDefinition> webPartsViaManager = cc.LoadQuery(limitedWPManager.WebParts.IncludeWithDefaultProperties(wp => wp.Id, wp => wp.WebPart.Title, wp => wp.WebPart.ZoneIndex, wp => wp.WebPart.IsClosed, wp => wp.WebPart.Hidden));

            cc.ExecuteQueryRetry();
            
            List<WebServiceWebPartProperties> webServiceWebPartEntities = LoadWebPartPropertiesFromWebServices(wikiPage.EnsureProperty(p => p.ServerRelativeUrl)); ;
            var pageUrl = page[Constants.FileRefField].ToString();

            foreach(var foundWebPart in webPartsViaManager)
            {
                var wsWpProps = webServiceWebPartEntities.SingleOrDefault(o => o.Id == foundWebPart.Id);
                if (wsWpProps != null && wsWpProps != default)
                {
                    var webPartPlaceholder = webPartsToRetrieve.SingleOrDefault(o => o.ControlId == wsWpProps.ControlId);
                    if (webPartPlaceholder != null && webPartPlaceholder != default)
                    {
                        webPartPlaceholder.WebPartDefinition = foundWebPart;
                    }
                }
            }

            // Load the web part XML for the web parts that do allow it
            //bool isDirty = false;
            foreach (var webPartToRetrieve in webPartsToRetrieve)
            {
                // Important to only process the web parts that did not return an error in the previous server call
                if (webPartToRetrieve.WebPartDefinition != null && (webPartToRetrieve.WebPartDefinition.ServerObjectIsNull.HasValue && webPartToRetrieve.WebPartDefinition.ServerObjectIsNull.Value == false))
                {
                    // If the web service call includes the export mode value then set the export options
                    var wsWp = webServiceWebPartEntities.FirstOrDefault(o => o.Id == webPartToRetrieve.WebPartDefinition.Id);
                    var wsExportMode = wsWp.Properties.FirstOrDefault(o => o.Key.Equals("exportmode", StringComparison.InvariantCultureIgnoreCase));
                    if (!string.IsNullOrEmpty(wsExportMode.Value) && wsExportMode.Value.Equals("all", StringComparison.InvariantCultureIgnoreCase))
                    {
                        var webPartXml = ExportWebPartXmlWorkaround(pageUrl, webPartToRetrieve.WebPartDefinition.Id.ToString());
                        webPartToRetrieve.WebPartXmlOnPremises = webPartXml;
                    }
                }
            }
           
            // Determine the web part type and store it in the web parts array
            foreach (var webPartToRetrieve in webPartsToRetrieve)
            {
                if (webPartToRetrieve.WebPartDefinition != null && (webPartToRetrieve.WebPartDefinition.ServerObjectIsNull.HasValue && webPartToRetrieve.WebPartDefinition.ServerObjectIsNull.Value == false))
                {
                    bool isExportable = false;

                    Dictionary<string, object> webPartProperties = null;

                    var wsWp = webServiceWebPartEntities.FirstOrDefault(o => o.Id == webPartToRetrieve.WebPartDefinition.Id);
                    webPartProperties = wsWp.PropertiesAsStringObjectDictionary();

                    // If the web service call includes the export mode value then set the export options
                    var wsExportMode = wsWp.Properties.FirstOrDefault(o => o.Key.Equals("exportmode", StringComparison.InvariantCultureIgnoreCase));
                    if (!string.IsNullOrEmpty(wsExportMode.Value) && wsExportMode.Value.ToString().Equals("all", StringComparison.InvariantCultureIgnoreCase))
                    {
                        isExportable = true;
                    }

                    if (!isExportable)
                    {
                        // Use different approach to determine type as we can't export the web part XML without indroducing a change
                        webPartToRetrieve.WebPartType = GetTypeFromProperties(webPartProperties, true);
                    }
                    else
                    {
                        webPartToRetrieve.WebPartType = GetType(webPartToRetrieve.WebPartXmlOnPremises);
                    }

                    webparts.Add(new WebPartEntity()
                    {
                        Title = webPartToRetrieve.WebPartDefinition.WebPart.Title,
                        Type = webPartToRetrieve.WebPartType,
                        Id = webPartToRetrieve.WebPartDefinition.Id,
                        ServerControlId = webPartToRetrieve.Id,
                        Row = webPartToRetrieve.Row,
                        Column = webPartToRetrieve.Column,
                        Order = webPartToRetrieve.Order,
                        ZoneId = "",
                        ZoneIndex = (uint)webPartToRetrieve.WebPartDefinition.WebPart.ZoneIndex,
                        IsClosed = webPartToRetrieve.WebPartDefinition.WebPart.IsClosed,
                        Hidden = webPartToRetrieve.WebPartDefinition.WebPart.Hidden,
                        Properties = Properties(webPartProperties, webPartToRetrieve.WebPartType, webPartToRetrieve.WebPartXmlOnPremises),
                    });
                }
            }
        }


        /// <summary>
        /// Stores text content as a fake web part
        /// </summary>
        /// <param name="wikiTextPartContent">Text to store</param>
        /// <param name="row">Row of the fake web part</param>
        /// <param name="col">Column of the fake web part</param>
        /// <param name="order">Order inside the row/column</param>
        /// <returns>A web part entity to add to the collection</returns>
        internal WebPartEntity CreateWikiTextPart(string wikiTextPartContent, int row, int col, int order)
        {
            Dictionary<string, string> properties = new Dictionary<string, string>();
            properties.Add("Text", wikiTextPartContent.Trim().Replace("\r\n", string.Empty));

            return new WebPartEntity()
            {
                Title = "WikiText",
                Type = "SharePointPnP.Modernization.WikiTextPart",
                Id = Guid.Empty,
                Row = row,
                Column = col,
                Order = order,
                Properties = properties,
            };
        }

        private void CleanHtml(IElement element, IHtmlDocument document)
        {
            foreach (var node in element.QuerySelectorAll("*").ToList())
            {
                if (node.ParentElement != null && IsUntransformableBlockElement(node))
                {
                    // create new div node and add all current children to it
                    var div = document.CreateElement("div");
                    foreach (var child in node.ChildNodes.ToList())
                    {
                        div.AppendChild(child);
                    }
                    // replace the unsupported node with the new div
                    node.ParentElement.ReplaceChild(div, node);
                }
            }
        }

        private bool IsUntransformableBlockElement(IElement element)
        {
            var tag = element.TagName.ToLower();
            if (tag == "article" ||
                tag == "address" ||
                tag == "aside" ||
                tag == "canvas" ||
                tag == "dd" ||
                tag == "dl" ||
                tag == "dt" ||
                tag == "fieldset" ||
                tag == "figcaption" ||
                tag == "figure" ||
                tag == "footer" ||
                tag == "form" ||
                tag == "header" ||
                //tag == "hr" || // will be replaced at in the html transformator
                tag == "main" ||
                tag == "nav" ||
                tag == "noscript" ||
                tag == "output" ||
                tag == "pre" ||
                tag == "section" ||
                tag == "tfoot" ||
                tag == "video" ||
                tag == "aside")
            {
                return true;
            }

            return false;
        }


        /// <summary>
        /// Does the tree of nodes somewhere contain a web part?
        /// </summary>
        /// <param name="element">Html content to analyze</param>
        /// <returns>True if it contains a web part</returns>
        private bool ContainsWebPart(IHtmlElement element)
        {
            var doc = parser.Parse(element.OuterHtml);
            var nodes = doc.All.Where(p => p.LocalName == "div");
            foreach (var node in nodes)
            {
                if (((node as IHtmlElement) != null) && (node as IHtmlElement).ClassList.Contains("ms-rte-wpbox"))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Strips the div holding the web part from the html
        /// </summary>
        /// <param name="element">Html element holding one or more web part divs</param>
        /// <returns>Cleaned html with a placeholder for the web part div</returns>
        private string StripWebPart(IHtmlElement element)
        {
            IElement copy = element.Clone(true) as IElement;
            var doc = parser.Parse(copy.OuterHtml);
            var nodes = doc.All.Where(p => p.LocalName == "div");
            if (nodes.Count() > 0)
            {
                foreach (var node in nodes.ToList())
                {
                    if (((node as IHtmlElement) != null) && (node as IHtmlElement).ClassList.Contains("ms-rte-wpbox"))
                    {
                        var newElement = doc.CreateTextNode(webPartMarkerString);
                        node.Parent.ReplaceChild(newElement, node);
                    }
                }

                if (doc.DocumentElement.Children[1].FirstElementChild != null &&
                    doc.DocumentElement.Children[1].FirstElementChild is IHtmlDivElement)
                {
                    return doc.DocumentElement.Children[1].FirstElementChild.InnerHtml;
                }
                else
                {
                    return doc.DocumentElement.Children[1].InnerHtml;
                }
            }
            else
            {
                return null;
            }
        }


        /// <summary>
        /// Get's the type of the web part by detecting if from the available properties
        /// </summary>
        /// <param name="properties">Web part properties to analyze</param>
        /// <returns>Type of the web part as fully qualified name</returns>
        public string GetTypeFromProperties(Dictionary<string, object> properties, bool isLegacy = false)
        {
            // Check for XSLTListView web part
            string[] xsltWebPart = new string[] { "ListUrl", "ListId", "Xsl", "JSLink", "ShowTimelineIfAvailable" };
            if (CheckWebPartProperties(xsltWebPart, properties))
            {
                return WebParts.XsltListView;
            }

            // Check for ListView web part
            string[] listWebPart = new string[] { "ListViewXml", "ListName", "ListId", "ViewContentTypeId", "PageType" };
            if (CheckWebPartProperties(listWebPart, properties))
            {
                return WebParts.ListView;
            }

            // check for Media web part
            string[] mediaWebPart = new string[] { "AutoPlay", "MediaSource", "Loop", "IsPreviewImageSourceOverridenForVideoSet", "PreviewImageSource" };
            if (CheckWebPartProperties(mediaWebPart, properties))
            {
                return WebParts.Media;
            }

            // check for SlideShow web part
            string[] slideShowWebPart = new string[] { "LibraryGuid", "Layout", "Speed", "ShowToolbar", "ViewGuid" };
            if (CheckWebPartProperties(slideShowWebPart, properties))
            {
                return WebParts.PictureLibrarySlideshow;
            }

            // check for Chart web part
            string[] chartWebPart = new string[] { "ConnectionPointEnabled", "ChartXml", "DataBindingsString", "DesignerChartTheme" };
            if (CheckWebPartProperties(chartWebPart, properties))
            {
                return WebParts.Chart;
            }

            // check for Site Members web part
            string[] membersWebPart = new string[] { "NumberLimit", "DisplayType", "MembershipGroupId", "Toolbar" };
            if (CheckWebPartProperties(membersWebPart, properties))
            {
                return WebParts.Members;
            }

            // check for Silverlight web part
            string[] silverlightWebPart = new string[] { "MinRuntimeVersion", "WindowlessMode", "CustomInitParameters", "Url", "ApplicationXml" };
            if (CheckWebPartProperties(silverlightWebPart, properties))
            {
                return WebParts.Silverlight;
            }

            // check for Add-in Part web part
            string[] addinPartWebPart = new string[] { "FeatureId", "ProductWebId", "ProductId" };
            if (CheckWebPartProperties(addinPartWebPart, properties))
            {
                return WebParts.Client;
            }

            if (isLegacy)
            {
                // Content Editor Web Part
                string[] contentEditorWebPart = new string[] { "Content", "ContentLink", "PartStorage" };
                if (CheckWebPartProperties(contentEditorWebPart, properties))
                {
                    return WebParts.ContentEditor;
                }

                // Image Viewer Web Part
                string[] imageViewerWebPart = new string[] { "ImageLink", "AlternativeText", "VerticalAlignment", "HorizontalAlignment" };
                if (CheckWebPartProperties(imageViewerWebPart, properties))
                {
                    return WebParts.Image;
                }

                // Title Bar 
                if(properties.ContainsKey("TypeName") && properties["TypeName"].ToString() == "Microsoft.SharePoint.WebPartPages.TitleBarWebPart")
                {
                    return WebParts.TitleBar;
                }


            }

            // check for Script Editor web part
            string[] scriptEditorWebPart = new string[] { "Content" };
            if (CheckWebPartProperties(scriptEditorWebPart, properties))
            {
                return WebParts.ScriptEditor;
            }

            // This needs to be last, but we still pages with sandbox user code web parts on them
            string[] sandboxWebPart = new string[] { "CatalogIconImageUrl", "AllowEdit", "TitleIconImageUrl", "ExportMode" };
            if (CheckWebPartProperties(sandboxWebPart, properties))
            {
                return WebParts.SPUserCode;
            }

            return "NonExportable_Unidentified";
        }

        private bool CheckWebPartProperties(string[] propertiesToCheck, Dictionary<string, object> properties)
        {
            bool isWebPart = true;
            foreach (var wpProp in propertiesToCheck)
            {
                if (!properties.ContainsKey(wpProp))
                {
                    isWebPart = false;
                    break;
                }
            }

            return isWebPart;
        }

        /// <summary>
        /// Checks the PageTransformation XML data to know which properties need to be kept for the given web part and collects their values
        /// </summary>
        /// <param name="properties">Properties collection retrieved when we loaded the web part</param>
        /// <param name="webPartType">Type of the web part</param>
        /// <param name="webPartXml">Web part XML</param>
        /// <returns>Collection of the requested property/value pairs</returns>
        public Dictionary<string, string> Properties(Dictionary<string, object> properties, string webPartType, string webPartXml)
        {
            Dictionary<string, string> propertiesToKeep = new Dictionary<string, string>();

            List<Property> propertiesToRetrieve = this.pageTransformation.BaseWebPart.Properties.ToList<Property>();

            //For older versions of SharePoint the type in the mapping would not match. Use the TypeShort Comparison. 
            var webPartProperties = this.pageTransformation.WebParts.Where(p => p.Type.GetTypeShort().Equals(webPartType.GetTypeShort(), StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
            if (webPartProperties != null && webPartProperties.Properties != null)
            {
                foreach (var p in webPartProperties.Properties.ToList<Property>())
                {
                    if (!propertiesToRetrieve.Contains(p))
                    {
                        propertiesToRetrieve.Add(p);
                    }
                }
            }

            if (string.IsNullOrEmpty(webPartXml))
            {
                if (webPartType == WebParts.Client)
                {
                    // Special case since we don't know upfront which properties are relevant here...so let's take them all
                    foreach (var prop in properties)
                    {
                        if (!propertiesToKeep.ContainsKey(prop.Key))
                        {
                            propertiesToKeep.Add(prop.Key, prop.Value != null ? prop.Value.ToString() : "");
                        }
                    }
                }
                else
                {
                    // Special case where we did not have export rights for the web part XML, assume this is a V3 web part
                    foreach (var property in propertiesToRetrieve)
                    {
                        if (!string.IsNullOrEmpty(property.Name) && properties.ContainsKey(property.Name))
                        {
                            if (!propertiesToKeep.ContainsKey(property.Name))
                            {
                                propertiesToKeep.Add(property.Name, properties[property.Name] != null ? properties[property.Name].ToString() : "");
                            }
                        }
                    }
                }
            }
            else
            {
                var xml = XElement.Parse(webPartXml);
                var xmlns = xml.XPathSelectElement("*").GetDefaultNamespace();
                if (xmlns.NamespaceName.Equals("http://schemas.microsoft.com/WebPart/v3", StringComparison.InvariantCultureIgnoreCase))
                {
                    if (webPartType == WebParts.Client)
                    {
                        // Special case since we don't know upfront which properties are relevant here...so let's take them all
                        foreach (var prop in properties)
                        {
                            if (!propertiesToKeep.ContainsKey(prop.Key))
                            {
                                propertiesToKeep.Add(prop.Key, prop.Value != null ? prop.Value.ToString() : "");
                            }
                        }
                    }
                    else
                    {
                        // the retrieved properties are sufficient
                        foreach (var property in propertiesToRetrieve)
                        {
                            if (!string.IsNullOrEmpty(property.Name) && properties.ContainsKey(property.Name))
                            {
                                if (!propertiesToKeep.ContainsKey(property.Name))
                                {
                                    propertiesToKeep.Add(property.Name, properties[property.Name] != null ? properties[property.Name].ToString() : "");
                                }
                            }
                        }
                    }
                }
                else if (xmlns.NamespaceName.Equals("http://schemas.microsoft.com/WebPart/v2", StringComparison.InvariantCultureIgnoreCase))
                {
                    foreach (var property in propertiesToRetrieve)
                    {
                        if (!string.IsNullOrEmpty(property.Name))
                        {
                            if (properties.ContainsKey(property.Name))
                            {
                                if (!propertiesToKeep.ContainsKey(property.Name))
                                {
                                    propertiesToKeep.Add(property.Name, properties[property.Name] != null ? properties[property.Name].ToString() : "");
                                }
                            }
                            else
                            {
                                // check XMl for property
                                var v2Element = xml.Descendants(xmlns + property.Name).FirstOrDefault();
                                if (v2Element != null)
                                {
                                    if (!propertiesToKeep.ContainsKey(property.Name))
                                    {
                                        propertiesToKeep.Add(property.Name, v2Element.Value);
                                    }
                                }

                                // Some properties do have their own namespace defined
                                if (webPartType == WebParts.SimpleForm && property.Name.Equals("Content", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    // Load using the http://schemas.microsoft.com/WebPart/v2/SimpleForm namespace
                                    XNamespace xmlcontentns = "http://schemas.microsoft.com/WebPart/v2/SimpleForm";
                                    v2Element = xml.Descendants(xmlcontentns + property.Name).FirstOrDefault();
                                    if (v2Element != null)
                                    {
                                        if (!propertiesToKeep.ContainsKey(property.Name))
                                        {
                                            propertiesToKeep.Add(property.Name, v2Element.Value);
                                        }
                                    }
                                }
                                else if (webPartType == WebParts.ContentEditor)
                                {
                                    if (property.Name.Equals("ContentLink", StringComparison.InvariantCultureIgnoreCase) ||
                                        property.Name.Equals("Content", StringComparison.InvariantCultureIgnoreCase) ||
                                        property.Name.Equals("PartStorage", StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        XNamespace xmlcontentns = "http://schemas.microsoft.com/WebPart/v2/ContentEditor";
                                        v2Element = xml.Descendants(xmlcontentns + property.Name).FirstOrDefault();
                                        if (v2Element != null)
                                        {
                                            if (!propertiesToKeep.ContainsKey(property.Name))
                                            {
                                                propertiesToKeep.Add(property.Name, v2Element.Value);
                                            }
                                        }
                                    }
                                }
                                else if (webPartType == WebParts.Xml)
                                {
                                    if (property.Name.Equals("XMLLink", StringComparison.InvariantCultureIgnoreCase) ||
                                        property.Name.Equals("XML", StringComparison.InvariantCultureIgnoreCase) ||
                                        property.Name.Equals("XSLLink", StringComparison.InvariantCultureIgnoreCase) ||
                                        property.Name.Equals("XSL", StringComparison.InvariantCultureIgnoreCase) ||
                                        property.Name.Equals("PartStorage", StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        XNamespace xmlcontentns = "http://schemas.microsoft.com/WebPart/v2/Xml";
                                        v2Element = xml.Descendants(xmlcontentns + property.Name).FirstOrDefault();
                                        if (v2Element != null)
                                        {
                                            if (!propertiesToKeep.ContainsKey(property.Name))
                                            {
                                                propertiesToKeep.Add(property.Name, v2Element.Value);
                                            }
                                        }
                                    }
                                }
                                else if (webPartType == WebParts.SiteDocuments)
                                {
                                    if (property.Name.Equals("UserControlledNavigation", StringComparison.InvariantCultureIgnoreCase) ||
                                        property.Name.Equals("ShowMemberships", StringComparison.InvariantCultureIgnoreCase) ||
                                        property.Name.Equals("UserTabs", StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        XNamespace xmlcontentns = "urn:schemas-microsoft-com:sharepoint:portal:sitedocumentswebpart";
                                        v2Element = xml.Descendants(xmlcontentns + property.Name).FirstOrDefault();
                                        if (v2Element != null)
                                        {
                                            if (!propertiesToKeep.ContainsKey(property.Name))
                                            {
                                                propertiesToKeep.Add(property.Name, v2Element.Value);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return propertiesToKeep;
        }

        /// <summary>
        /// Call SharePoint Web Services to extract web part properties not exposed by CSOM
        /// </summary>
        /// <returns></returns>
        internal Tuple<string,string> ExtractWebPartDocumentViaWebServicesFromPage(string fullDocumentUrl)
        {
            try
            {
                LogInfo(LogStrings.CallingWebServicesToExtractWebPartsFromPage, LogStrings.Heading_ContentTransform);
                
                string webUrl = cc.Web.GetUrl();
                string webServiceUrl = webUrl + "/_vti_bin/WebPartPages.asmx";

                StringBuilder soapEnvelope = new StringBuilder();

                soapEnvelope.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                soapEnvelope.Append("<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">");

                soapEnvelope.Append(String.Format(
                 "<soap:Body>" +
                     "<GetWebPartPage xmlns=\"http://microsoft.com/sharepoint/webpartpages\">" +
                         "<documentName>{0}</documentName>" +
                         "<behavior>Version3</behavior>" +
                     "</GetWebPartPage>" +
                 "</soap:Body>"
                 , fullDocumentUrl));

                soapEnvelope.Append("</soap:Envelope>");

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(webServiceUrl);
                request.Credentials = cc.Credentials;
                request.Method = "POST";
                request.ContentType = "text/xml; charset=\"utf-8\"";
                request.Accept = "text/xml";
                request.Headers.Add("SOAPAction", "\"http://microsoft.com/sharepoint/webpartpages/GetWebPartPage\"");

                using (System.IO.Stream stream = request.GetRequestStream())
                {
                    using (System.IO.StreamWriter writer = new System.IO.StreamWriter(stream))
                    {
                        writer.Write(soapEnvelope.ToString());
                    }
                }

                var response = request.GetResponse();
                using (var dataStream = response.GetResponseStream())
                {
                    XmlDocument xDoc = new XmlDocument();
                    xDoc.Load(dataStream);

                    if (xDoc.DocumentElement != null && xDoc.DocumentElement.InnerText.Length > 0)
                    {
                        var webPartPageContents = xDoc.DocumentElement.InnerText;
                        //Remove the junk from the result
                        var tag = "<HasByteOrderMark/>";
                        var marker = webPartPageContents.IndexOf(tag);
                        var partDocument = string.Empty;
                        if (marker > -1)
                        {
                            partDocument = webPartPageContents.Substring(marker).Replace(tag, "");
                        }

                        return new Tuple<string, string>(webPartPageContents, partDocument);
                    }
                }
            }
            catch (WebException ex)
            {
                LogError(LogStrings.Error_CallingWebServicesToExtractWebPartsFromPage,LogStrings.Heading_ContentTransform, ex);
            }

            return new Tuple<string, string>(string.Empty, string.Empty);
        }
                
        /// <summary>
        /// Exports Web Part XML via an older workround
        /// </summary>
        /// <param name="pageUrl"></param>
        /// <param name="webPartGuid"></param>
        /// <returns></returns>
        public string ExportWebPartXmlWorkaround(string pageUrl, string webPartGuid)
        {
            // Issue hints and Credit: 
            //      https://blog.mastykarz.nl/export-web-parts-csom/ 
            //      https://sharepoint.stackexchange.com/questions/30865/missing-export-option-for-sharepoint-2010-webparts
            //      https://github.com/SharePoint/PnP-Sites-Core/pull/908/files

            try
            {
                LogInfo(LogStrings.RetreivingExportWebPartXmlWorkaround, LogStrings.Heading_ContentTransform);
                string webPartXml = String.Empty;
                string serverRelativeUrl = cc.Web.EnsureProperty(w => w.ServerRelativeUrl);
                var uri = new Uri(cc.Site.Url);

                var fullWebUrl = $"{uri.Scheme}://{uri.Host}:{uri.Port}{serverRelativeUrl}";
                var fullPageUrl = $"{uri.Scheme}://{uri.Host}:{uri.Port}{pageUrl}";

                if (!fullWebUrl.EndsWith("/"))
                {
                    fullWebUrl = fullWebUrl + "/";
                }

                string webServiceUrl = string.Format("{0}_vti_bin/exportwp.aspx?pageurl={1}&guidstring={2}", fullWebUrl, fullPageUrl, webPartGuid);

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(webServiceUrl); //hack to force webpart zones to render
                request.Credentials = cc.Credentials;
                request.Method = "GET";

                var response = request.GetResponse();
                using (var dataStream = response.GetResponseStream())
                {
                    XmlDocument xDoc = new XmlDocument();
                    xDoc.Load(dataStream);

                    if (xDoc.DocumentElement != null && xDoc.DocumentElement.OuterXml.Length > 0)
                    {
                        webPartXml = xDoc.DocumentElement.OuterXml;

                        // Not sure what causes the web parts to switch from singular to multiple
                        if(xDoc.DocumentElement.LocalName == "webParts")
                        {
                            webPartXml = xDoc.DocumentElement.InnerXml;
                        }

                        return webPartXml;
                    }
                }
            }
            catch (WebException ex)
            {
                LogError(LogStrings.Error_ExportWebPartXmlWorkaround, LogStrings.Heading_ContentTransform, ex);
            }

            return string.Empty;
        }

        /// <summary>
        /// Loads and Parses Web Part Page from Web Services
        /// </summary>
        /// <param name="fullUrl"></param>
        public List<WebServiceWebPartEntity> LoadPublishingPageFromWebServices(string fullUrl)
        {
            var webParts = new List<WebServiceWebPartEntity>();
            var wsWebParts = ExtractWebPartDocumentViaWebServicesFromPage(fullUrl);

            if (!string.IsNullOrEmpty(wsWebParts.Item2))
            {
                var doc = parser.Parse(wsWebParts.Item2);

                List<Tuple<string, string>> prefixesAndNameSpaces = ExtractWebPartPrefixesFromNamespaces(doc);
                List<Tuple<string, string>> possibleWebPartsUsed = new List<Tuple<string, string>>();

                prefixesAndNameSpaces.ForEach(p =>
                {
                    var possibleParts = WebParts.GetListOfWebParts(p.Item2);
                    foreach (var part in possibleParts)
                    {
                        var webPartName = part.Substring(0, part.IndexOf(",")).Replace($"{p.Item2}.", "");
                        possibleWebPartsUsed.Add(new Tuple<string, string>(webPartName, part));
                    }
                });

                var xmlBlock = wsWebParts.Item2;
                var tag = "</head>";
                var marker = xmlBlock.IndexOf(tag);
                if (marker > -1)
                {
                    xmlBlock = xmlBlock.Substring(marker).Replace(tag, "");
                }

                // Clean prefixes
                xmlBlock = xmlBlock.Replace("__designer:", "Designer");
                foreach(var prefix in prefixesAndNameSpaces)
                {
                    xmlBlock = xmlBlock.Replace($"{prefix.Item1}:", "");
                }

                if (!string.IsNullOrEmpty(xmlBlock))
                {
                    XmlDocument xDoc = new XmlDocument();
                    xDoc.LoadXml(xmlBlock);

                    if (xDoc.DocumentElement != null)
                    {
                        var childNodes = xDoc.SelectNodes("//ZoneTemplate/*");
                        foreach (XmlNode node in childNodes)
                        {
                            XmlNode nodeToExtractProperties = node;
                            WebServiceWebPartEntity webPart = new WebServiceWebPartEntity();

                            //This should only find one match
                            var matchWebPart = possibleWebPartsUsed.FirstOrDefault(o => o.Item1.ToUpper() == nodeToExtractProperties.LocalName.ToUpper());

                            if (matchWebPart != default)
                            {
                                webPart.Type = matchWebPart.Item2;
                            }

                            var wpId = nodeToExtractProperties.Attributes.GetNamedItem("__WebPartId");
                            webPart.Id = Guid.Parse(wpId?.Value);

                            // In the case of Content Editor web parts
                            if (node.HasChildNodes && node.FirstChild.LocalName == "WebPart")
                            {
                                // Some web parts store properties as child nodes
                                nodeToExtractProperties = node.FirstChild;

                                foreach (XmlNode wpChildNodes in nodeToExtractProperties.ChildNodes)
                                {
                                    webPart.Properties.Add(wpChildNodes.LocalName, wpChildNodes.InnerText);
                                }
                            }
                            else
                            {
                                // Some web parts store properties by attributes
                                foreach (XmlAttribute attr in nodeToExtractProperties.Attributes)
                                {
                                    webPart.Properties.Add(attr.Name, attr.Value);
                                }
                            }

                            webParts.Add(webPart);
                        }
                    }
                }
            }

            return webParts;
        }

        /// <summary>
        /// Gets the tag prefixes from the document
        /// </summary>
        /// <param name="webPartPage"></param>
        /// <returns></returns>
        internal List<Tuple<string, string>> ExtractWebPartPrefixesFromNamespaces(IHtmlDocument webPartPage)
        {
            var tagPrefixes = new List<Tuple<string, string>>();

            Regex regex = new Regex("&lt;%@(.*?)%&gt;", RegexOptions.IgnoreCase | RegexOptions.Multiline);
            var aspxHeader = webPartPage.All.Where(o => o.TagName == "HTML").FirstOrDefault();
            var results = regex.Matches(aspxHeader?.InnerHtml);

            StringBuilder blockHtml = new StringBuilder();
            foreach (var match in results)
            {
                var matchString = match.ToString().Replace("&lt;%@ ", "<").Replace("%&gt;", " />");
                blockHtml.AppendLine(matchString);
            }

            var fullBlock = blockHtml.ToString();
            using (var subDocument = this.parser.Parse(fullBlock))
            {
                var registers = subDocument.All.Where(o => o.TagName == "REGISTER");

                foreach (var register in registers)
                {
                    var prefix = register.GetAttribute("Tagprefix");
                    var nameSpace = register.GetAttribute("Namespace");
                    var className = nameSpace.InferClassNameFromNameSpace();
                    tagPrefixes.Add(new Tuple<string, string>(prefix, nameSpace));
                    tagPrefixes.Add(new Tuple<string, string>(className, nameSpace));
                }

            }

            return tagPrefixes;
        }

        /// <summary>
        /// Loads Web Part Properties from web services
        /// </summary>
        /// <param name="pageUrl">Server Relative Page Url</param>
        /// <returns></returns>
        public List<WebServiceWebPartProperties> LoadWebPartPropertiesFromWebServices(string pageUrl)

        {
            // This may contain references to multiple web part
            var webPartProperties = new List<WebServiceWebPartProperties>();
            var wsWebPartProps = ExtractWebPartPropertiesViaWebServicesFromPage(pageUrl);

            if (!string.IsNullOrEmpty(wsWebPartProps))
            {
                try
                {
                    XmlDocument xDoc = new XmlDocument();
                    xDoc.LoadXml(wsWebPartProps);
                    var namespaceManager = new XmlNamespaceManager(xDoc.NameTable);
                    namespaceManager.AddNamespace("soap", "http://schemas.xmlsoap.org/soap/envelope/");
                    namespaceManager.AddNamespace("wpp", "http://microsoft.com/sharepoint/webpartpages");
                    namespaceManager.AddNamespace("wpv3", "http://schemas.microsoft.com/WebPart/v3");
                    namespaceManager.AddNamespace("wpv2", "http://schemas.microsoft.com/WebPart/v2");

                    if (xDoc.DocumentElement != null && xDoc.DocumentElement.InnerXml.Length > 0)
                    {
                        var xPath = "//soap:Body/wpp:GetWebPartProperties2Response/wpp:GetWebPartProperties2Result/wpp:WebParts/wpv2:WebPart";
                        var webParts = xDoc.SelectNodes(xPath, namespaceManager);

                        foreach (XmlNode webPartNode in webParts)
                        {
                            WebServiceWebPartProperties wsWebPartPropertiesEntity = new WebServiceWebPartProperties();

                            // Get Parent Node for Web Part ID
                            var parentNodeIdAttribute = webPartNode.Attributes.GetNamedItem("ID");
                            wsWebPartPropertiesEntity.Id = Guid.Parse(parentNodeIdAttribute.Value);

                            // Get the control ID from properties
                            var controlIDNode = webPartNode.SelectSingleNode("wpv2:ID", namespaceManager);
                            wsWebPartPropertiesEntity.ControlId = controlIDNode?.InnerText; 

                            // Get the type from a child node
                            var typeNode = webPartNode.SelectSingleNode("wpv2:TypeName", namespaceManager);
                            wsWebPartPropertiesEntity.Type = typeNode?.InnerText;

                            // Get the type from a child node
                            var zoneIDNode = webPartNode.SelectSingleNode("wpv2:ZoneID", namespaceManager);
                            wsWebPartPropertiesEntity.ZoneId = zoneIDNode?.InnerText;
                            
                            //Get the properties
                            var propertyNodes = webPartNode.ChildNodes;

                            foreach (XmlNode propertyNode in propertyNodes)
                            {
                                wsWebPartPropertiesEntity.Properties.Add(propertyNode.LocalName, propertyNode.InnerText);
                            }

                            webPartProperties.Add(wsWebPartPropertiesEntity);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Catch exceptions during processing of Web Service Responses
                }
            }

            return webPartProperties;
        }

        /// <summary>
        /// Call SharePoint Web Services to extract web part properties not exposed by CSOM
        /// </summary>
        /// <returns></returns>
        internal string ExtractWebPartPropertiesViaWebServicesFromPage(string pageUrl)
        {
            try
            {
                LogInfo(LogStrings.CallingWebServicesToExtractWebPartPropertiesFromPage, LogStrings.Heading_ContentTransform);
                string webPartProperties = String.Empty;
                string webUrl = cc.Web.GetUrl();
                string webServiceUrl = webUrl + "/_vti_bin/WebPartPages.asmx";

                StringBuilder soapEnvelope = new StringBuilder();

                soapEnvelope.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                soapEnvelope.Append("<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">");

                soapEnvelope.Append(String.Format(
                 "<soap:Body>" +
                     "<GetWebPartProperties2 xmlns=\"http://microsoft.com/sharepoint/webpartpages\">" +
                         "<pageUrl>{0}</pageUrl>" +
                         "<storage>Shared</storage>" +
                         "<behavior>Version3</behavior>" +
                     "</GetWebPartProperties2>" +
                 "</soap:Body>"
                 , pageUrl));

                soapEnvelope.Append("</soap:Envelope>");

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(webServiceUrl); //hack to force webpart zones to render
                request.Credentials = cc.Credentials;
                request.Method = "POST";
                request.ContentType = "text/xml; charset=\"utf-8\"";
                request.Accept = "text/xml";
                request.Headers.Add("SOAPAction", "\"http://microsoft.com/sharepoint/webpartpages/GetWebPartProperties2\"");

                using (System.IO.Stream stream = request.GetRequestStream())
                {
                    using (System.IO.StreamWriter writer = new System.IO.StreamWriter(stream))
                    {
                        writer.Write(soapEnvelope.ToString());
                    }
                }

                var response = request.GetResponse();
                using (var dataStream = response.GetResponseStream())
                {
                    XmlDocument xDoc = new XmlDocument();
                    xDoc.Load(dataStream);

                    if (xDoc.DocumentElement != null && xDoc.DocumentElement.InnerXml.Length > 0)
                    {
                        webPartProperties = xDoc.DocumentElement.InnerXml;

                        return webPartProperties;
                    }
                }
            }
            catch (WebException ex)
            {
                LogError(LogStrings.Error_ExtractWebPartPropertiesViaWebServicesFromPage, LogStrings.Heading_ContentTransform, ex);
            }

            return string.Empty;
        }

    }
}
