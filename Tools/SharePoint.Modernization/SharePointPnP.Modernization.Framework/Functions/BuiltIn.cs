﻿using AngleSharp;
using AngleSharp.Parser.Html;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Pages;
using SharePointPnP.Modernization.Framework.Telemetry;
using SharePointPnP.Modernization.Framework.Transform;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace SharePointPnP.Modernization.Framework.Functions
{
    /// <summary>
    /// Set of native, builtin, functions
    /// </summary>
    public partial class BuiltIn : FunctionsBase
    {
        private ClientContext sourceClientContext;
        private ClientContext targetClientContext;
        private ClientSidePage clientSidePage;
        private BaseTransformationInformation baseTransformationInformation;
        private UrlTransformator urlTransformator;

        #region Construction
        /// <summary>
        /// Instantiates the base builtin function library
        /// </summary>
        /// <param name="pageClientContext">ClientContext object for the site holding the page being transformed</param>
        /// <param name="sourceClientContext">The ClientContext for the source </param>
        /// <param name="clientSidePage">Reference to the client side page</param>
        public BuiltIn(BaseTransformationInformation baseTransformationInformation, ClientContext pageClientContext, ClientContext sourceClientContext = null, ClientSidePage clientSidePage = null, IList<ILogObserver> logObservers = null) : base(pageClientContext)
        {
            if (logObservers != null)
            {
                foreach (var observer in logObservers)
                {
                    base.RegisterObserver(observer);
                }
            }

            // This is an optional property, in cross site transfer the two contexts would be different.
            this.sourceClientContext = sourceClientContext;
            this.targetClientContext = pageClientContext;
            this.clientSidePage = clientSidePage;
            this.baseTransformationInformation = baseTransformationInformation;
            this.urlTransformator = new UrlTransformator(baseTransformationInformation, this.sourceClientContext, this.targetClientContext, base.RegisteredLogObservers);
        }
        #endregion

        // All functions return either a single string, boolean or a Dictionary<string,string> with key value pairs. 
        // Allowed input parameter types are string, int, bool, DateTime and Guid

        #region Generic functions
        /// <summary>
        /// Html encodes a string
        /// </summary>
        /// <param name="text">Text to html encode</param>
        /// <returns>Html encoded string</returns>
        [FunctionDocumentation(Description = "Returns the html encoded value of this string.",
                               Example = "{EncodedText} = HtmlEncode({Text})")]
        [InputDocumentation(Name = "{Text}", Description = "Text to html encode")]
        [OutputDocumentation(Name = "{EncodedText}", Description = "Html encoded text")]
        public string HtmlEncode(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return System.Web.HttpUtility.HtmlEncode(text);
        }

        /// <summary>
        /// Html encodes string for inclusion in JSON
        /// </summary>
        /// <param name="text">Text to html encode</param>
        /// <returns>Html encoded string for inclusion in JSON</returns>
        [FunctionDocumentation(Description = "Returns the json html encoded value of this string.",
                               Example = "{JsonEncodedText} = HtmlEncodeForJson({Text})")]
        [InputDocumentation(Name = "{Text}", Description = "Text to html encode for inclusion in json")]
        [OutputDocumentation(Name = "{JsonEncodedText}", Description = "Html encoded text for inclusion in json file")]
        public string HtmlEncodeForJson(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return System.Web.HttpUtility.HtmlEncode(text).Replace("&quot;", @"\&quot;").Replace(":", "&#58;").Replace("@", "%40");
        }

        /// <summary>
        /// Return true
        /// </summary>
        /// <returns>True</returns>
        [FunctionDocumentation(Description = "Simply returns the string true.",
                               Example = "{UsePlaceHolders} = ReturnTrue()")]
        [OutputDocumentation(Name = "{UsePlaceHolders}", Description = "Value true")]
        public string ReturnTrue()
        {
            return "true";
        }

        /// <summary>
        /// Return false
        /// </summary>
        /// <returns>False</returns>
        [FunctionDocumentation(Description = "Simply returns the string false.",
                               Example = "{UsePlaceHolders} = ReturnFalse()")]
        [OutputDocumentation(Name = "{UsePlaceHolders}", Description = "Value false")]
        public string ReturnFalse()
        {
            return "false";
        }

        /// <summary>
        /// Transforms the incoming path into a server relative path
        /// </summary>
        /// <param name="path">Path to transform</param>
        /// <returns>Server relative path</returns>
        [FunctionDocumentation(Description = "Transforms the incoming path into a server relative path.",
                               Example = "{ServerRelativePath} = ReturnServerRelativePath({Path})")]
        [InputDocumentation(Name = "{Path}", Description = "Path to transform")]
        [OutputDocumentation(Name = "{ServerRelativePath}", Description = "Server relative path")]
        public string ReturnServerRelativePath(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return "";
            }

            var hostUri = new Uri(this.clientContext.Web.GetUrl());
            string host = $"{hostUri.Scheme}://{hostUri.DnsSafeHost}";

            return path.Replace(host, "");
        }

        /// <summary>
        /// Returns the filename of the given path
        /// </summary>
        /// <param name="path"></param>
        /// <returns>File name</returns>
        [FunctionDocumentation(Description = "Returns the filename of the given path.",
                               Example = "{FileName} = ReturnFileName({Path})")]
        [InputDocumentation(Name = "{Path}", Description = "Path to analyze")]
        [OutputDocumentation(Name = "{FileName}", Description = "File name with extension from the given path")]
        public string ReturnFileName(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return "";
            }

            return Path.GetFileName(path);
        }

        [FunctionDocumentation(Description = "Concatenates 2 strings.",
                               Example = "{CompleteString} = Concatenate({String1},{String2})")]
        [InputDocumentation(Name = "{String1}", Description = "First string")]
        [InputDocumentation(Name = "{String2}", Description = "Second string")]
        [OutputDocumentation(Name = "{CompleteString}", Description = "Concatenation of the passed strings")]
        public string Concatenate(string string1, string string2)
        {
            if (string1 == null)
            {
                string1 = "";
            }
            if (string2 == null)
            {
                string2 = "";
            }

            return string1 + string2;
        }

        /// <summary>
        /// Returns an the (static) string provided as input
        /// </summary>
        /// <returns>String provided as input</returns>
        [FunctionDocumentation(Description = "Returns an the (static) string provided as input",
                               Example = "StaticString('static string')")]
        [InputDocumentation(Name = "'static string'", Description = "Static input string")]
        [OutputDocumentation(Name = "return value", Description = "String provided as input")]
        public string StaticString(string staticString)
        {
            return staticString;
        }
        #endregion

        #region Text functions
        /// <summary>
        /// Selector to allow to embed a spacer instead of an empty text
        /// </summary>
        /// <param name="text">Text to evaluate</param>
        /// <returns>Text if text needs to be inserted, Spacer if text was empty and you want a spacer</returns>
        [SelectorDocumentation(Description = "Allows for option to include a spacer for empty text wiki text parts.",
                               Example = "TextSelector({CleanedText})")]
        [InputDocumentation(Name = "{CleanedText}", Description = "Client side text part compliant html (cleaned via TextCleanup function)")]
        [OutputDocumentation(Name = "Text", Description = "Will be output if the provided wiki text was not considered empty")]
        [OutputDocumentation(Name = "Spacer", Description = "Will be output if the provided wiki text was considered empty")]
        public string TextSelector(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "Text";
            }

            var empty = new HtmlTransformator().IsEmptyParagraph(text);

            if (empty)
            {
                return "Spacer";
            }
            else
            {
                return "Text";
            }
        }

        /// <summary>
        /// Wiki html rewrite to work in RTE
        /// </summary>
        /// <param name="text">Wiki html to rewrite</param>
        /// <returns>Html that's compatible with RTE</returns>
        [FunctionDocumentation(Description = "Rewrites wiki page html to be compliant with the html supported by the client side text part.",
                               Example = "{CleanedText} = TextCleanup({Text},{UsePlaceHolders})")]
        [InputDocumentation(Name = "{Text}", Description = "Original wiki html content")]
        [InputDocumentation(Name = "{UsePlaceHolders}", Description = "Parameter indicating if placeholders must be included for unsupported img/iframe elements inside wiki html")]
        [OutputDocumentation(Name = "{CleanedText}", Description = "Html compliant with client side text part")]
        public string TextCleanup(string text, string usePlaceHolders)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            // Rewrite url's if needed
            if (!this.baseTransformationInformation.SkipUrlRewrite)
            {
                text = this.urlTransformator.Transform(text);
            }

            bool usePlaceHolder = true;

            bool.TryParse(usePlaceHolders, out usePlaceHolder);

            return new HtmlTransformator().Transform(text, usePlaceHolder);
        }

        /// <summary>
        /// Checks if the provided html contains JavaScript
        /// </summary>
        /// <param name="content">Html content to check</param>
        /// <returns>True is the html contains script, false otherwise</returns>
        [FunctionDocumentation(Description = "Checks if the provided html contains JavaScript",
                               Example = "{HasScript} = ContainsScript({Text})")]
        [InputDocumentation(Name = "{Text}", Description = "Html content to check")]
        [OutputDocumentation(Name = "{HasScript}", Description = "True is the html contains script, false otherwise")]
        public bool ContainsScript(string content)
        {
            if (string.IsNullOrEmpty(content))
            {
                return false;
            }

            HtmlParser parser = new HtmlParser(new HtmlParserOptions() { IsEmbedded = true }, Configuration.Default.WithDefaultLoader().WithCss());

            try
            {
                var doc = parser.Parse(content);
                // Script information
                var scriptTags = doc.All.Where(p => p.LocalName == "script");
                if (scriptTags.Count() > 0)
                {
                    return true;
                }
            }
            catch { }

            return false;
        }
        #endregion

        #region List functions, used by XsltListViewWebPart
        /// <summary>
        /// Selector that returns the base type of the list as input for selecting the correct mapping
        /// </summary>
        /// <param name="listId">Id of the list</param>
        /// <returns>Mapping to be used for the given list</returns>
        [SelectorDocumentation(Description = "Analyzes a list and returns the list base type.",
                               Example = "ListSelectorListLibrary({ListId})")]
        [InputDocumentation(Name = "{ListId}", Description = "Guid of the list to use")]
        [InputDocumentation(Name = "{ViewXml}", Description = "Definition of the selected view")]
        [OutputDocumentation(Name = "Library", Description = "The list is a document library")]
        [OutputDocumentation(Name = "List", Description = "The list is a document list")]
        [OutputDocumentation(Name = "Issue", Description = "The list is an issue list")]
        [OutputDocumentation(Name = "TaskList", Description = "The list is an task list")]
        [OutputDocumentation(Name = "DiscussionBoard", Description = "The list is a discussion board")]
        [OutputDocumentation(Name = "Survey", Description = "The list is a survey")]
        [OutputDocumentation(Name = "Undefined", Description = "The list base type is undefined")]
        public string ListSelectorListLibrary(Guid listId, string viewXml)
        {
            if (listId == Guid.Empty)
            {
                return "";
            }
            else
            {
                if (!string.IsNullOrEmpty(viewXml))
                {
                    if (viewXml.IndexOf("Type=\"CALENDAR\"", StringComparison.InvariantCultureIgnoreCase) > -1)
                    {
                        return "Calendar";
                    }
                }

                var list = this.clientContext.Web.GetListById(listId);
                list.EnsureProperties(p => p.BaseType, p => p.BaseTemplate);

                // "Detailed" inspection based on template
                if (list.BaseTemplate == (int)ListTemplateType.Tasks || list.BaseTemplate == (int)ListTemplateType.TasksWithTimelineAndHierarchy)
                {
                    return "TaskList";
                }
                else if (list.BaseTemplate == (int)ListTemplateType.DiscussionBoard)
                {
                    return "DiscussionBoard";
                }

                // "Generic" inspection based on type
                if (list.BaseType == BaseType.DocumentLibrary)
                {
                    return "Library";
                }
                else if (list.BaseType == BaseType.GenericList)
                {
                    return "List";
                }
                else if (list.BaseType == BaseType.Issue)
                {
                    return "Issue";
                }
                else if (list.BaseType == BaseType.DiscussionBoard)
                {
                    return "DiscussionBoard";
                }
                else if (list.BaseType == BaseType.Survey)
                {
                    return "Survey";
                }

                return "Undefined";
            }
        }

        /// <summary>
        /// Returns the cross site collection save list id.
        /// </summary>
        /// <param name="listId">Id of the list</param>
        /// <returns>Cross site collection safe list id</returns>
        [FunctionDocumentation(Description = "Returns the cross site collection save list id.",
                               Example = "{ListId} = ListCrossSiteCheck({ListId})")]
        [InputDocumentation(Name = "{ListId}", Description = "Guid of the list to use")]
        [OutputDocumentation(Name = "{ListId}", Description = "Cross site collection safe list id")]
        public string ListCrossSiteCheck(Guid listId)
        {
            if (listId == Guid.Empty)
            {
                return "";
            }
            else
            {
                var sourceList = this.sourceClientContext.Web.GetListById(listId);
                sourceList.EnsureProperty(p=>p.Title);

                List targetlist = null;
                try
                {
                    targetlist = this.clientContext.Web.GetListByTitle(sourceList.Title);
                    targetlist.EnsureProperty(p => p.Id);
                }
                catch (Exception ex)
                {
                    throw new NotAvailableAtTargetException($"List with id {listId} and Title {sourceList.Title} is not available in the target site collection. This web part will be skipped.", ex);
                }

                return targetlist.Id.ToString();
            }
        }

        /// <summary>
        /// Function that returns the server relative url of the given list
        /// </summary>
        /// <param name="listId">Id of the list</param>
        /// <returns>Server relative url of the list</returns>
        [FunctionDocumentation(Description = "Returns the server relative url of a list.",
                               Example = "{ListServerRelativeUrl} = ListAddServerRelativeUrl({ListId})")]
        [InputDocumentation(Name = "{ListId}", Description = "Guid of the list to use")]
        [OutputDocumentation(Name = "{ListServerRelativeUrl}", Description = "Server relative url of the list")]
        public string ListAddServerRelativeUrl(Guid listId)
        {
            if (listId == Guid.Empty)
            {
                return "";
            }
            else
            {
                var list = this.clientContext.Web.GetListById(listId);
                list.EnsureProperty(p => p.RootFolder).EnsureProperty(p => p.ServerRelativeUrl);
                return list.RootFolder.ServerRelativeUrl;
            }
        }

        /// <summary>
        /// Function that returns the web relative url of the given list
        /// </summary>
        /// <param name="listId">Id of the list</param>
        /// <returns>Web relative url of the list</returns>
        [FunctionDocumentation(Description = "Returns the web relative url of a list.",
                               Example = "{ListWebRelativeUrl} = ListAddWebRelativeUrl({ListId})")]
        [InputDocumentation(Name = "{ListId}", Description = "Guid of the list to use")]
        [OutputDocumentation(Name = "{ListWebRelativeUrl}", Description = "Web relative url of the list")]
        public string ListAddWebRelativeUrl(Guid listId)
        {
            if (listId == Guid.Empty)
            {
                return "";
            }
            else
            {
                var list = this.clientContext.Web.GetListById(listId);
                list.EnsureProperty(p => p.RootFolder).EnsureProperty(p => p.ServerRelativeUrl);
                this.clientContext.Web.EnsureProperty(p => p.ServerRelativeUrl);

                // For lists in the rootweb of the root site collection of the tenant the replacement is not needed
                if (!String.IsNullOrEmpty(this.clientContext.Web.ServerRelativeUrl.TrimEnd('/')))
                {
                    return list.RootFolder.ServerRelativeUrl.Replace(this.clientContext.Web.ServerRelativeUrl.TrimEnd('/'), "");
                }
                else
                {
                    return list.RootFolder.ServerRelativeUrl;
                }
            }
        }

        /// <summary>
        /// Checks if an XSLTListView web part has a hidden toolbar
        /// </summary>
        /// <param name="xmlDefinition">XmlDefinition attribute of the XSLTListViewWebPart</param>
        /// <returns>Boolean indicating if the toolbar should be hidden</returns>
        [FunctionDocumentation(Description = "Checks if an XSLTListView web part has a hidden toolbar.",
                               Example = "{HideToolBar} = ListHideToolBar({XmlDefinition})")]
        [InputDocumentation(Name = "{XmlDefinition}", Description = "XmlDefinition attribute of the XSLTListViewWebPart")]
        [OutputDocumentation(Name = "{HideToolBar}", Description = "Boolean indicating if the toolbar should be hidden")]
        public bool ListHideToolBar(string xmlDefinition)
        {
            if (string.IsNullOrEmpty(xmlDefinition))
            {
                return false;
            }

            // Get the "identifying" elements from the webpart view xml definition
            var webPartViewElement = XElement.Parse(xmlDefinition);

            var toolBar = webPartViewElement.Descendants("Toolbar").FirstOrDefault();
            if (toolBar != null)
            {
                string toolBarType = toolBar.Attribute("Type") != null ? toolBar.Attribute("Type").Value : null;

                if (!string.IsNullOrEmpty(toolBarType))
                {
                    if (toolBarType.Equals("None", StringComparison.InvariantCultureIgnoreCase))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Tries to find the id of the view used to configure the web part
        /// </summary>
        /// <param name="listId">Id of the list</param>
        /// <param name="xmlDefinition">Webpart view definition</param>
        /// <returns>Id of the detected view if found or otherwise the id of the default list view</returns>
        [FunctionDocumentation(Description = "Detects the list view id that was used by the webpart by mapping the web part xmldefinition to the list views. If no view found the list default view id is returned.",
                               Example = "{ListViewId} = ListDetectUsedView({ListId},{XmlDefinition})")]
        [InputDocumentation(Name = "{ListId}", Description = "Guid of the list to analyze")]
        [InputDocumentation(Name = "{XmlDefinition}", Description = "XmlDefinition attribute of the XSLTListViewWebPart")]
        [OutputDocumentation(Name = "{ListViewId}", Description = "Id of the view to be used")]
        public string ListDetectUsedView(Guid listId, string xmlDefinition)
        {
            if (listId == Guid.Empty || string.IsNullOrEmpty(xmlDefinition))
            {
                return "";
            }

            // Grab the list and the needed properties
            var list = this.clientContext.Web.GetListById(listId);
            list.EnsureProperties(l => l.DefaultView, l => l.Views.Include(v => v.Hidden, v => v.Id, v => v.ListViewXml));

            // Get the "identifying" elements from the webpart view xml definition
            var webPartViewElement = XElement.Parse(xmlDefinition);

            // Analyze the views in the list to determine a possible mapping
            foreach (var view in list.Views.AsEnumerable().Where(view => !view.Hidden && view.ListViewXml != null))
            {
                var viewElement = XElement.Parse(view.ListViewXml);

                // Compare Query
                if (webPartViewElement.Descendants("Query").FirstOrDefault() != null && viewElement.Descendants("Query").FirstOrDefault() != null)
                {
                    var equalNodes = XmlComparer.AreEqual(webPartViewElement.Descendants("Query").FirstOrDefault(), viewElement.Descendants("Query").FirstOrDefault());
                    if (!equalNodes.Success)
                    {
                        continue;
                    }
                }
                else
                {
                    if (!(webPartViewElement.Descendants("Query").FirstOrDefault() == null && viewElement.Descendants("Query").FirstOrDefault() != null))
                    {
                        continue;
                    }
                }

                // Compare viewFields
                if (webPartViewElement.Descendants("ViewFields").FirstOrDefault() != null && viewElement.Descendants("ViewFields").FirstOrDefault() != null)
                {
                    var equalNodes = XmlComparer.AreEqual(webPartViewElement.Descendants("ViewFields").FirstOrDefault(), viewElement.Descendants("ViewFields").FirstOrDefault());
                    if (!equalNodes.Success)
                    {
                        continue;
                    }
                }
                else
                {
                    if (!(webPartViewElement.Descendants("ViewFields").FirstOrDefault() == null && viewElement.Descendants("ViewFields").FirstOrDefault() != null))
                    {
                        continue;
                    }
                }

                // Compare RowLimit
                if (webPartViewElement.Descendants("RowLimit").FirstOrDefault() != null && viewElement.Descendants("RowLimit").FirstOrDefault() != null)
                {
                    var equalNodes = XmlComparer.AreEqual(webPartViewElement.Descendants("RowLimit").FirstOrDefault(), viewElement.Descendants("RowLimit").FirstOrDefault());
                    if (!equalNodes.Success)
                    {
                        continue;
                    }
                }
                else
                {
                    if (!(webPartViewElement.Descendants("RowLimit").FirstOrDefault() == null && viewElement.Descendants("RowLimit").FirstOrDefault() != null))
                    {
                        continue;
                    }
                }

                // Yeah, we're still here so we found the matching view!
                return view.Id.ToString();
            }

            // No matching view found, so proceed with the default view
            return list.DefaultView.Id.ToString();
        }

        #endregion

        #region Image functions
        /// <summary>
        /// Does return image properties based on given server relative image path
        /// </summary>
        /// <param name="serverRelativeImagePath">Server relative path of the image</param>
        /// <returns>A set of image properties</returns>
        [FunctionDocumentation(Description = "Does lookup a file based on the given server relative path and return needed properties of the file. Returns null if file was not found.",
                               Example = "ImageLookup({ServerRelativeFileName})")]
        [InputDocumentation(Name = "{ServerRelativeFileName}", Description = "Server relative file name of the image")]
        [OutputDocumentation(Name = "{ImageListId}", Description = "Id of the list holding the file")]
        [OutputDocumentation(Name = "{ImageUniqueId}", Description = "UniqueId of the file")]
        public Dictionary<string, string> ImageLookup(string serverRelativeImagePath)
        {

            bool stop = false;
            if (string.IsNullOrEmpty(serverRelativeImagePath))
            {
                stop = true;
            }

            this.clientContext.Web.EnsureProperty(p => p.ServerRelativeUrl);

            // Check if this url is pointing to content living in this site
            if (!stop && !serverRelativeImagePath.StartsWith(this.clientContext.Web.ServerRelativeUrl, StringComparison.InvariantCultureIgnoreCase))
            {
                // We're not looking up the image, providing the server relative path to the modern Image web part is sufficient
                stop = true;
            }

            Dictionary<string, string> results = new Dictionary<string, string>();

            if (stop)
            {
                results.Add("ImageListId", "");
                results.Add("ImageUniqueId", "");
                return results;
            }

            try
            {
                var pageHeaderImage = this.clientContext.Web.GetFileByServerRelativeUrl(serverRelativeImagePath);
                this.clientContext.Load(pageHeaderImage, p => p.UniqueId, p => p.ListId);
                this.clientContext.ExecuteQueryRetry();

                results.Add("ImageListId", pageHeaderImage.ListId.ToString());
                results.Add("ImageUniqueId", pageHeaderImage.UniqueId.ToString());
                return results;
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                {
                    // Provided image was not found, should not happen
                    return null;
                }
                else
                {
                    throw;
                }
            }
        }

        /// <summary>
        /// Copy the asset to target site in cross site transformation
        /// </summary>
        /// <param name="imageLink"></param>
        [FunctionDocumentation(Description = "Transforms the incoming path into a server relative path. If the page is located on another page the asset is transferred and url updated. Any failures keep to the original value.",
            Example = "{ServerRelativeFileName} = ReturnCrossSiteRelativePath({ImageLink})")]
        [InputDocumentation(Name = "{ImageLink}", Description = "Original value for the image link")]
        [OutputDocumentation(Name = "{ServerRelativeFileName}", Description = "New target location for the asset if transferred.")]
        public string ReturnCrossSiteRelativePath(string imageLink)
        {
            // Defaults to the orignal operation
            var serverRelativeAssetFileName = ReturnServerRelativePath(imageLink);

            try
            {
                string pageFileName = null;

                if (this.clientSidePage != null && !string.IsNullOrEmpty(this.clientSidePage.PageTitle))
                {
                    pageFileName = this.clientSidePage.PageTitle;
                }
                else
                {
                    // deduct based upon filename in url
                    pageFileName = Path.GetFileNameWithoutExtension(serverRelativeAssetFileName);
                }

                AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, base.clientContext, base.RegisteredLogObservers);

                var newAssetLocation = assetTransfer.TransferAsset(serverRelativeAssetFileName, pageFileName);

                return newAssetLocation;

            }
            catch (Exception ex)
            {
                LogError(LogStrings.Error_ReturnCrossSiteRelativePath, LogStrings.Heading_BuiltInFunctions, ex);
            }
                       
            return serverRelativeAssetFileName;
        }
        #endregion

        #region First party web parts hosted on classic pages
        /// <summary>
        /// Extracts the client side web part properties so they can be reused
        /// </summary>
        /// <param name="clientSideWebPartHtml">Html defining the client side web part hosted on a classic page</param>
        /// <returns>Client side web part properties ready for reuse</returns>
        [FunctionDocumentation(Description = "Extracts the client side web part properties so they can be reused.",
                               Example = "{JsonProperties} = ExtractWebpartProperties({ClientSideWebPartData})")]
        [InputDocumentation(Name = "{ClientSideWebPartData}", Description = "Web part data defining the client side web part configuration")]
        [OutputDocumentation(Name = "{JsonProperties}", Description = "Json properties to configure the client side web part")]
        public string ExtractWebpartProperties(string clientSideWebPartHtml)
        {
            if (string.IsNullOrEmpty(clientSideWebPartHtml))
            {
                return "{}";
            }

            HtmlParser parser = new HtmlParser(new HtmlParserOptions() { IsEmbedded = true });
            using (var document = parser.Parse(clientSideWebPartHtml))
            {
                return document.Body.FirstElementChild.GetAttribute("data-sp-webpartdata");
            }
        }
        #endregion

        #region DocumentEmbed functions
        [FunctionDocumentation(Description = "Does lookup a file based on the given server relative path and return needed properties of the file. Returns null if file was not found.",
                               Example = "DocumentEmbedLookup({ServerRelativeFileName})")]
        [InputDocumentation(Name = "{ServerRelativeFileName}", Description = "Server relative file name")]
        [OutputDocumentation(Name = "{DocumentListId}", Description = "Id of the list holding the file")]
        [OutputDocumentation(Name = "{DocumentUniqueId}", Description = "UniqueId of the file")]
        [OutputDocumentation(Name = "{DocumentAuthor}", Description = "User principal name of the document author")]
        [OutputDocumentation(Name = "{DocumentAuthorName}", Description = "Name of the file author")]
        public Dictionary<string, string> DocumentEmbedLookup(string serverRelativeUrl)
        {
            Dictionary<string, string> results = new Dictionary<string, string>();
            if (string.IsNullOrEmpty(serverRelativeUrl))
            {
                results.Add("DocumentWeb", "");
                results.Add("DocumentListId", "");
                results.Add("DocumentUniqueId", "");
                results.Add("DocumentAuthor", "");
                results.Add("DocumentAuthorName", "");
                return results;
            }

            // Assume document lives in current web
            ClientContext contextToUse = this.clientContext;

            this.clientContext.Web.EnsureProperty(p => p.ServerRelativeUrl);
            if (!serverRelativeUrl.StartsWith(this.clientContext.Web.ServerRelativeUrl, StringComparison.InvariantCultureIgnoreCase))
            {
                try
                {
                    // 
                    Uri hostUri = new Uri(this.clientContext.Web.GetUrl());

                    // Find the web url hosting the content file
                    var webUrlResult = Web.GetWebUrlFromPageUrl(this.clientContext, $"{hostUri.Scheme}://{hostUri.DnsSafeHost}{serverRelativeUrl}");
                    this.clientContext.ExecuteQueryRetry();

                    contextToUse = this.clientContext.Clone(webUrlResult.Value);
                }
                catch (Exception ex)
                {
                    LogError(LogStrings.Error_DocumentEmbedLookup, LogStrings.Heading_BuiltInFunctions, ex);

                    results.Add("DocumentWeb", "");
                    results.Add("DocumentListId", "");
                    results.Add("DocumentUniqueId", "");
                    results.Add("DocumentAuthor", "");
                    results.Add("DocumentAuthorName", "");
                    return results;
                }
            }

            try
            {
                var document = contextToUse.Web.GetFileByServerRelativeUrl(serverRelativeUrl);
                contextToUse.Load(document, p => p.UniqueId, p => p.ListId, p => p.Author);
                contextToUse.Load(contextToUse.Web, p => p.ServerRelativeUrl);
                contextToUse.ExecuteQueryRetry();

                string[] authorParts = document.Author.LoginName.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

                results.Add("DocumentWeb", contextToUse.Web.ServerRelativeUrl);
                results.Add("DocumentListId", document.ListId.ToString());
                results.Add("DocumentUniqueId", document.UniqueId.ToString());
                results.Add("DocumentAuthor", authorParts.Length == 3 ? authorParts[2] : "");
                results.Add("DocumentAuthorName", document.Author.Title);

                return results;
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                {
                    // provided file is not retrievable...we're eating the exception this file not be used in the target web part
                    LogError(LogStrings.Error_DocumentEmbedLookupFileNotRetrievable, LogStrings.Heading_BuiltInFunctions, ex);
                    return null;
                }
                else
                {
                    LogError(LogStrings.Error_DocumentEmbedLookup, LogStrings.Heading_BuiltInFunctions, ex);
                    throw;
                }
            }
        }
        #endregion

        #region Content Embed functions
        [SelectorDocumentation(Description = "Analyzes sourcetype and return recommended mapping.",
                               Example = "ContentEmbedSelectorSourceType({SourceType})")]
        [InputDocumentation(Name = "{SourceType}", Description = "Sourcetype of the viewed page in pageviewerwebpart")]
        [OutputDocumentation(Name = "WebPage", Description = "The embedded content is a page")]
        [OutputDocumentation(Name = "ServerFolderOrFile", Description = "The embedded content points to a server folder or file")]
        public string ContentEmbedSelectorSourceType(string sourceType)
        {
            if (sourceType == "4")
            {
                return "WebPage";
            }

            return "ServerFolderOrFile";
        }

        [SelectorDocumentation(Description = "Content editor can be transformed in various ways depending on whether a link was used, what file type was used, if script is used or not...",
                               Example = "ContentEmbedSelectorContentLink({ContentLink}, {Content}, {FileContents}, {UseCommunityScriptEditor})")]
        [InputDocumentation(Name = "{ContentLink}", Description = "Link value if set")]
        [InputDocumentation(Name = "{Content}", Description = "Content embedded inside the web part")]
        [InputDocumentation(Name = "{FileContents}", Description = "Text content of the file. Return empty string if file was not found")]
        [InputDocumentation(Name = "{UseCommunityScriptEditor}", Description = "The UseCommunityScriptEditor mapping property provided via the PageTransformationInformation instance")]
        [OutputDocumentation(Name = "Link", Description = "If the link was not empty and it was an aspx file")]
        [OutputDocumentation(Name = "NonASPXLink", Description = "If the link was not empty and it was not an aspx file but the file contents did contain JavaScript")]
        [OutputDocumentation(Name = "NonASPXLinkNoScript", Description = "If the link was not empty and it was not an aspx file and the contents did not contain JavaScript")]
        [OutputDocumentation(Name = "NonASPXUseCommunityScriptEditor", Description = "Use the community script editor to host the content")]
        [OutputDocumentation(Name = "Content", Description = "If no link was specified but content was embedded and it contains JavaScript")]
        [OutputDocumentation(Name = "ContentNoScript", Description = "If no link was specified and the embedded content and it does not contain JavaScript")]
        [OutputDocumentation(Name = "ContentUseCommunityScriptEditor", Description = "Use the community script editor to host the content")]
        public string ContentEmbedSelectorContentLink(string contentLink, string embeddedContent, string fileContent, string useCommunityScriptEditor)
        {
            bool.TryParse(useCommunityScriptEditor, out bool useCommunityScriptEditorBool);

            if (!string.IsNullOrEmpty(contentLink))
            {
                if (contentLink.ToLower().EndsWith(".aspx"))
                {
                    return "Link";
                }
                else
                {
                    if (!ContainsScript(fileContent))
                    {
                        return "NonASPXLinkNoScript";
                    }
                    else
                    {
                        if (useCommunityScriptEditorBool)
                        {
                            return "NonASPXUseCommunityScriptEditor";
                        }
                        else
                        {
                            return "NonASPXLink";
                        }
                    }
                }
            }
            else
            {
                if (!ContainsScript(embeddedContent))
                {
                    return "ContentNoScript";
                }
                else
                {
                    if (useCommunityScriptEditorBool)
                    {
                        return "ContentUseCommunityScriptEditor";
                    }
                    else
                    {
                        return "Content";
                    }
                }
            }
        }

        /// <summary>
        /// Throws an exception when link to .aspx file.
        /// </summary>
        /// <param name="listId">Link value if set</param>
        /// <returns>Unused variable</returns>
        [FunctionDocumentation(Description = "Throws an exception when link to .aspx file.",
                               Example = "{Temp} = ContentEmbedCrossSiteCheck({ContentLink})")]
        [InputDocumentation(Name = "{ContentLink}", Description = "Link value if set")]
        [OutputDocumentation(Name = "{Temp}", Description = "Unused variable")]
        public string ContentEmbedCrossSiteCheck(string contentLink)
        {

            if (! IsCrossSiteTransfer() || string.IsNullOrEmpty(contentLink))
            {
                return "";
            }
            else
            {
                if (contentLink.ToLower().EndsWith(".aspx"))
                {
                    throw new NotAvailableAtTargetException($"ASPX Page with link {contentLink} is not available in the target site collection. This web part will be skipped.");
                }
            }

            return "";
        }

        /// <summary>
        /// Loads contents of a file as a string.
        /// </summary>
        /// <param name="contentLink">Server relative url to the file to load</param>
        /// <returns>Text content of the file. Return empty string if file was not found</returns>
        [FunctionDocumentation(Description = "Loads contents of a file as a string.",
                               Example = "{FileContents} = LoadContentFromFile({ContentLink})")]
        [InputDocumentation(Name = "{ContentLink}", Description = "Server relative url to the file to load")]
        [OutputDocumentation(Name = "{FileContents}", Description = "Text content of the file. Return empty string if file was not found")]

        public string LoadContentFromFile(string contentLink)
        {
            if (string.IsNullOrEmpty(contentLink) || Path.GetExtension(contentLink).Equals(".aspx", StringComparison.InvariantCultureIgnoreCase))
            {
                return "";
            }

            this.clientContext.Web.EnsureProperty(p => p.ServerRelativeUrl);
            if (!contentLink.StartsWith(this.clientContext.Web.ServerRelativeUrl, StringComparison.InvariantCultureIgnoreCase))
            {
                try
                {
                    // Content editor does allow a web part on a sub web to point to a file in the rootweb...Pointing to files outside of the current site collection is not allowed
                    Uri hostUri = new Uri(this.clientContext.Web.GetUrl());

                    // Find the web url hosting the content file
                    var webUrlResult = Web.GetWebUrlFromPageUrl(this.clientContext, $"{hostUri.Scheme}://{hostUri.DnsSafeHost}{contentLink}");
                    this.clientContext.ExecuteQueryRetry();

                    using (var cc = this.clientContext.Clone(webUrlResult.Value))
                    {
                        return cc.Web.GetFileAsString(contentLink);
                    }
                }
                catch (Exception ex)
                {
                    LogError(LogStrings.Error_LoadContentFromFile, LogStrings.Heading_BuiltInFunctions, ex);
                    return "";
                }
            }
            else
            {
                try
                {
                    return this.sourceClientContext.Web.GetFileAsString(contentLink);
                }
                catch (ServerException ex)
                {
                    if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                    {
                        // Provided html was not found, should not happen but if it happens we're not stopping the transformation
                        LogError(LogStrings.Error_LoadContentFromFileContentLink, LogStrings.Heading_BuiltInFunctions, ex);
                        return "";
                    }
                    else
                    {
                        LogError(LogStrings.Error_LoadContentFromFileContentLink, LogStrings.Heading_BuiltInFunctions, ex);
                        throw;
                    }
                }
            }
        }
        #endregion

        #region HighlightedContent functions
        /// <summary>
        /// Maps the user documents web part data into a properties collection and supporting serverProcessedContent nodes for the content rollup (= Highlighted Content) web part
        /// </summary>
        /// <param name="dataProviderJson"></param>
        /// <param name="selectedPropertiesJson"></param>
        /// <param name="resultsPerPage"></param>
        /// <param name="renderTemplateId"></param>
        /// <returns>A properties collection and supporting serverProcessedContent nodes for the content rollup (= Highlighted Content) web part</returns>
        [FunctionDocumentation(Description = "Maps the user documents web part data into a properties collection and supporting serverProcessedContent nodes for the content rollup (= Highlighted Content) web part",
                                   Example = "SiteDocumentsToHighlightedContentProperties()")]
        [OutputDocumentation(Name = "JsonProperties", Description = "Properties collection for the contentrollup (= Highlighted Content) web part")]
        [OutputDocumentation(Name = "SearchablePlainTexts", Description = "SearchablePlainTexts nodes to be added in the serverProcessedContent node")]
        [OutputDocumentation(Name = "Links", Description = "Links nodes to be added in the serverProcessedContent node")]
        [OutputDocumentation(Name = "ImageSources", Description = "ImageSources nodes to be added in the serverProcessedContent node")]
        public Dictionary<string, string> UserDocumentsToHighlightedContentProperties()
        {
            Dictionary<string, string> results = new Dictionary<string, string>();

            ContentByQuerySearchTransformator cqs = new ContentByQuerySearchTransformator(this.clientContext);
            var res = cqs.TransformUserDocuments();

            // Output the calculated properties so then can be used in the mapping
            results.Add("JsonProperties", res.Properties);
            results.Add("SearchablePlainTexts", res.SearchablePlainTexts);
            results.Add("Links", res.Links);
            results.Add("ImageSources", res.ImageSources);

            return results;
        }

        /// <summary>
        /// Maps content by search web part data into a properties collection and supporting serverProcessedContent nodes for the content rollup (= Highlighted Content) web part
        /// </summary>
        /// <param name="dataProviderJson"></param>
        /// <param name="selectedPropertiesJson"></param>
        /// <param name="resultsPerPage"></param>
        /// <param name="renderTemplateId"></param>
        /// <returns>A properties collection and supporting serverProcessedContent nodes for the content rollup (= Highlighted Content) web part</returns>
        [FunctionDocumentation(Description = "Maps content by search web part data into a properties collection and supporting serverProcessedContent nodes for the content rollup (= Highlighted Content) web part",
                                   Example = "ContentBySearchToHighlightedContentProperties({DataProviderJSON}, {SelectedPropertiesJson}, {ResultsPerPage}, {RenderTemplateId})")]
        [InputDocumentation(Name = "{DataProviderJson}", Description = "")]
        [InputDocumentation(Name = "{SelectedPropertiesJson}", Description = "")]
        [InputDocumentation(Name = "{ResultsPerPage}", Description = "")]
        [InputDocumentation(Name = "{RenderTemplateId}", Description = "")]
        [OutputDocumentation(Name = "JsonProperties", Description = "Properties collection for the contentrollup (= Highlighted Content) web part")]
        [OutputDocumentation(Name = "SearchablePlainTexts", Description = "SearchablePlainTexts nodes to be added in the serverProcessedContent node")]
        [OutputDocumentation(Name = "Links", Description = "Links nodes to be added in the serverProcessedContent node")]
        [OutputDocumentation(Name = "ImageSources", Description = "ImageSources nodes to be added in the serverProcessedContent node")]
        public Dictionary<string, string> ContentBySearchToHighlightedContentProperties(string dataProviderJson, string selectedPropertiesJson, int resultsPerPage, string renderTemplateId)
        {
            Dictionary<string, string> results = new Dictionary<string, string>();

            ContentBySearch cbs = new ContentBySearch()
            {
                DataProviderJson = dataProviderJson,
                SelectedPropertiesJson = selectedPropertiesJson,
                ResultsPerPage = resultsPerPage,
                RenderTemplateId = renderTemplateId
            };

            ContentByQuerySearchTransformator cqs = new ContentByQuerySearchTransformator(this.clientContext);
            var res = cqs.TransformContentBySearchWebPartToHighlightedContent(cbs);

            // Output the calculated properties so then can be used in the mapping
            results.Add("JsonProperties", res.Properties);
            results.Add("SearchablePlainTexts", res.SearchablePlainTexts);
            results.Add("Links", res.Links);
            results.Add("ImageSources", res.ImageSources);

            return results;
        }

        /// <summary>
        /// Maps content by query web part data into a properties collection for the contentrollup (= Highlighted Content) web part
        /// </summary>
        /// <param name="webUrl"></param>
        /// <param name="listGuid"></param>
        /// <param name="listName"></param>
        /// <param name="serverTemplate"></param>
        /// <param name="contentTypeBeginsWithId"></param>
        /// <param name="filterField1"></param>
        /// <param name="filter1ChainingOperator"></param>
        /// <param name="filterField1Value"></param>
        /// <param name="filterOperator1"></param>
        /// <param name="filterField2"></param>
        /// <param name="filter2ChainingOperator"></param>
        /// <param name="filterField2Value"></param>
        /// <param name="filterOperator2"></param>
        /// <param name="filterField3"></param>
        /// <param name="filterField3Value"></param>
        /// <param name="filterOperator3"></param>
        /// <param name="sortBy"></param>
        /// <param name="sortByDirection"></param>
        /// <param name="groupBy"></param>
        /// <param name="groupByDirection"></param>
        /// <param name="itemLimit"></param>
        /// <param name="displayColumns"></param>
        /// <param name="dataMappings"></param>
        /// <returns>A properties collection and supporting serverProcessedContent nodes for the content rollup (= Highlighted Content) web part</returns>
        [FunctionDocumentation(Description = "Maps content by query web part data into a properties collection and supporting serverProcessedContent nodes for the content rollup (= Highlighted Content) web part",
                               Example = "ContentByQueryToHighlightedContentProperties({WebUrl},{ListGuid},{ListName},{ServerTemplate},{ContentTypeBeginsWithId},{FilterField1},{Filter1ChainingOperator},{FilterDisplayValue1},{FilterOperator1},{FilterField2},{Filter2ChainingOperator},{FilterDisplayValue2},{FilterOperator2},{FilterField3},{FilterDisplayValue3},{FilterOperator3},{SortBy},{SortByDirection},{GroupBy},{GroupByDirection},{ItemLimit},{DisplayColumns},{DataMappings})")]
        [InputDocumentation(Name = "{WebUrl}", Description = "")]
        [InputDocumentation(Name = "{ListGuid}", Description = "")]
        [InputDocumentation(Name = "{ListName}", Description = "")]
        [InputDocumentation(Name = "{ServerTemplate}", Description = "")]
        [InputDocumentation(Name = "{ContentTypeBeginsWithId}", Description = "")]
        [InputDocumentation(Name = "{FilterField1}", Description = "")]
        [InputDocumentation(Name = "{Filter1ChainingOperator}", Description = "")]
        [InputDocumentation(Name = "{FilterField1Value}", Description = "")]
        [InputDocumentation(Name = "{FilterOperator1}", Description = "")]
        [InputDocumentation(Name = "{FilterField2}", Description = "")]
        [InputDocumentation(Name = "{Filter2ChainingOperator}", Description = "")]
        [InputDocumentation(Name = "{FilterField2Value}", Description = "")]
        [InputDocumentation(Name = "{FilterOperator2}", Description = "")]
        [InputDocumentation(Name = "{FilterField3}", Description = "")]
        [InputDocumentation(Name = "{FilterField3Value}", Description = "")]
        [InputDocumentation(Name = "{FilterOperator3}", Description = "")]
        [InputDocumentation(Name = "{SortBy}", Description = "")]
        [InputDocumentation(Name = "{SortByDirection}", Description = "")]
        [InputDocumentation(Name = "{GroupBy}", Description = "")]
        [InputDocumentation(Name = "{GroupByDirection}", Description = "")]
        [InputDocumentation(Name = "{ItemLimit}", Description = "")]
        [InputDocumentation(Name = "{DisplayColumns}", Description = "")]
        [InputDocumentation(Name = "{DataMappings}", Description = "")]
        [OutputDocumentation(Name = "JsonProperties", Description = "Properties collection for the contentrollup (= Highlighted Content) web part")]
        [OutputDocumentation(Name = "SearchablePlainTexts", Description = "SearchablePlainTexts nodes to be added in the serverProcessedContent node")]
        [OutputDocumentation(Name = "Links", Description = "Links nodes to be added in the serverProcessedContent node")]
        [OutputDocumentation(Name = "ImageSources", Description = "ImageSources nodes to be added in the serverProcessedContent node")]
        public Dictionary<string, string> ContentByQueryToHighlightedContentProperties(string webUrl, string listGuid, string listName, string serverTemplate, string contentTypeBeginsWithId,
                                                                                       string filterField1, string filter1ChainingOperator, string filterField1Value, string filterOperator1,
                                                                                       string filterField2, string filter2ChainingOperator, string filterField2Value, string filterOperator2,
                                                                                       string filterField3, string filterField3Value, string filterOperator3,
                                                                                       string sortBy, string sortByDirection, string groupBy, string groupByDirection, string itemLimit, int displayColumns,
                                                                                       string dataMappings)
        {
            Dictionary<string, string> results = new Dictionary<string, string>();

            ContentByQuery cbq = new ContentByQuery()
            {
                WebUrl = webUrl,
                ListGuid = listGuid,
                ListName = listName,
                ServerTemplate = serverTemplate,
                ContentTypeBeginsWithId = contentTypeBeginsWithId,

                FilterField1 = filterField1,
                Filter1ChainingOperator = (FilterChainingOperator)Enum.Parse(typeof(FilterChainingOperator), filter1ChainingOperator, true),
                FilterField1Value = filterField1Value,
                FilterOperator1 = (FilterFieldQueryOperator)Enum.Parse(typeof(FilterFieldQueryOperator), filterOperator1, true),
                FilterField2 = filterField2,
                Filter2ChainingOperator = (FilterChainingOperator)Enum.Parse(typeof(FilterChainingOperator), filter2ChainingOperator, true),
                FilterField2Value = filterField2Value,
                FilterOperator2 = (FilterFieldQueryOperator)Enum.Parse(typeof(FilterFieldQueryOperator), filterOperator2, true),
                FilterField3 = filterField3,
                FilterField3Value = filterField3Value,
                FilterOperator3 = (FilterFieldQueryOperator)Enum.Parse(typeof(FilterFieldQueryOperator), filterOperator3, true),

                SortBy = sortBy,
                SortByDirection = (SortDirection)Enum.Parse(typeof(SortDirection), sortByDirection, true),
                GroupBy = groupBy,
                GroupByDirection = (SortDirection)Enum.Parse(typeof(SortDirection), groupByDirection, true),

                ItemLimit = Convert.ToInt32(itemLimit),
                DisplayColumns = displayColumns,

                DataMappings = dataMappings
            };

            ContentByQuerySearchTransformator cqs = new ContentByQuerySearchTransformator(this.clientContext);
            var res = cqs.TransformContentByQueryWebPartToHighlightedContent(cbq);

            // Output the calculated properties so then can be used in the mapping
            results.Add("JsonProperties", res.Properties);
            results.Add("SearchablePlainTexts", res.SearchablePlainTexts);
            results.Add("Links", res.Links);
            results.Add("ImageSources", res.ImageSources);

            return results;
        }

        [SelectorDocumentation(Description = "Analyzes a list and returns if the list can be transformed.",
                               Example = "ContentByQuerySelector({ListGuid},{ListName})")]
        [InputDocumentation(Name = "{ListGuid}", Description = "Guid of the list used by the CBQ web part")]
        [InputDocumentation(Name = "{ListName}", Description = "Name of the list used by the CBQ web part")]
        [OutputDocumentation(Name = "Default", Description = "Transform the list")]
        [OutputDocumentation(Name = "NoTransformation", Description = "Don't transform the list")]
        public string ContentByQuerySelector(string listGuid, string listName)
        {

            // Scoped to list?
            Guid.TryParse(listGuid, out Guid listId);

            if (!string.IsNullOrEmpty(listName) || listId != Guid.Empty)
            {
                // Scope to list
                List list = null;
                if (listId != Guid.Empty)
                {
                    list = this.clientContext.Web.GetListById(listId);
                }
                else
                {
                    list = this.clientContext.Web.GetListByTitle(listName);
                }

                this.clientContext.Load(list, p => p.BaseType);
                this.clientContext.ExecuteQueryRetry();

                if (list.BaseType != BaseType.DocumentLibrary)
                {
                    return "NoTransformation";
                }
            }

            return "Default";
        }
        #endregion

        #region SummaryLink functions
        /// <summary>
        /// Uses the SummaryLinksToQuickLinks mapping property provided via the PageTransformationInformation instance to determine the mapping
        /// </summary>
        /// <param name="useQuickLinks">The SummaryLinksToQuickLinks mapping property provided via the PageTransformationInformation instance</param>
        /// <returns>Whether to transform via the QuickLinks web part or via Text</returns>
        [SelectorDocumentation(Description = "Uses the SummaryLinksToQuickLinks mapping property provided via the PageTransformationInformation instance to determine the mapping",
                               Example = "SummaryLinkSelector({SummaryLinksToQuickLinks})")]
        [InputDocumentation(Name = "{SummaryLinksToQuickLinks}", Description = "The SummaryLinksToQuickLinks mapping property provided via the PageTransformationInformation instance")]
        [OutputDocumentation(Name = "UseQuickLinks", Description = "Transform to the QuickLinks web part")]
        [OutputDocumentation(Name = "UseText", Description = "Transform to the formatted text")]
        public string SummaryLinkSelector(string useQuickLinks)
        {
            if (bool.TryParse(useQuickLinks, out bool useQuickLinksBool))
            {
                if (useQuickLinksBool)
                {
                    return "UseQuickLinks";
                }
            }

            return "UseText";
        }

        /// <summary>
        /// Rewrites summarylinks web part html to be compliant with the html supported by the client side text part
        /// </summary>
        /// <param name="text">Original wiki html content</param>
        /// <returns>Html compliant with client side text part</returns>
        [FunctionDocumentation(Description = "Rewrites summarylinks web part html to be compliant with the html supported by the client side text part.",
                       Example = "{CleanedText} = TextCleanUpSummaryLinks({Text})")]
        [InputDocumentation(Name = "{Text}", Description = "Original wiki html content")]
        [OutputDocumentation(Name = "{CleanedText}", Description = "Html compliant with client side text part")]
        public string TextCleanUpSummaryLinks(string text, int chromeType, string title)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            // Add header for all but the chrometype = none and chrometype = border only
            string webPartTitle = null;
            if (chromeType != 2 && chromeType != 4)
            {
                if (!string.IsNullOrEmpty(title))
                {
                    webPartTitle = title;
                }
            }

            // Rewrite url's if needed
            if (!this.baseTransformationInformation.SkipUrlRewrite)
            {
                text = this.urlTransformator.Transform(text);
            }

            var summaryLinksHtmlTransformator = new SummaryLinksHtmlTransformator
            {
                WebPartTitle = webPartTitle
            };

            return summaryLinksHtmlTransformator.Transform(text, false);
        }

        /// <summary>
        /// Maps summarylinks web part data into a properties collection and supporting serverProcessedContent nodes for the quicklinks web part
        /// </summary>
        /// <param name="text">Original wiki html content</param>
        /// <returns>Properties collection for the quicklinks web part</returns>
        [FunctionDocumentation(Description = "Maps summarylinks web part data into a properties collection and supporting serverProcessedContent nodes for the quicklinks web part",
                               Example = "SummaryLinksToQuickLinksProperties({Text},{QuickLinksJsonProperties})")]
        [InputDocumentation(Name = "{Text}", Description = "Original wiki html content")]
        [InputDocumentation(Name = "{QuickLinksJsonProperties}", Description = "QuickLinks JSON properties blob (optional)")]
        [OutputDocumentation(Name = "JsonProperties", Description = "Properties collection for the quicklinks web part")]
        [OutputDocumentation(Name = "SearchablePlainTexts", Description = "SearchablePlainTexts nodes to be added in the serverProcessedContent node")]
        [OutputDocumentation(Name = "Links", Description = "Links nodes to be added in the serverProcessedContent node")]
        [OutputDocumentation(Name = "ImageSources", Description = "ImageSources nodes to be added in the serverProcessedContent node")]
        public Dictionary<string, string> SummaryLinksToQuickLinksProperties(string text, string quickLinksJsonProperties = "")
        {
            Dictionary<string, string> results = new Dictionary<string, string>();

            var links = new SummaryLinksHtmlTransformator().GetLinks(text);

            if (IsCrossSiteTransfer())
            {
                var clientSidePage = this.clientSidePage.PageTitle;
                AssetTransfer assetTransfer = new AssetTransfer(sourceClientContext, base.clientContext, base.RegisteredLogObservers);

                foreach (var link in links)
                {
                    // preview images
                    if (!string.IsNullOrEmpty(link.ImageUrl))
                    {
                        var serverRelativeAssetFileName = ReturnServerRelativePath(link.ImageUrl);
                        var newAssetLocation = assetTransfer.TransferAsset(serverRelativeAssetFileName, clientSidePage);
                        link.ImageUrl = newAssetLocation;
                    }

                    // urls
                    if (!string.IsNullOrEmpty(link.Url))
                    {
                        var serverRelativeAssetFileName = ReturnServerRelativePath(link.Url);
                        var newAssetLocation = assetTransfer.TransferAsset(serverRelativeAssetFileName, clientSidePage);
                        link.Url = newAssetLocation;
                    }
                }
            }

            // Rewrite url's if needed
            if (!this.baseTransformationInformation.SkipUrlRewrite)
            {
                foreach(var link in links)
                {
                    link.Url = this.urlTransformator.Transform(link.Url);
                }
            }

            QuickLinksTransformator qlt = new QuickLinksTransformator(this.clientContext, base.RegisteredLogObservers);
            var res = qlt.Transform(links, quickLinksJsonProperties);

            // Output the calculated properties so then can be used in the mapping
            results.Add("JsonProperties", res.Properties);
            results.Add("SearchablePlainTexts", res.SearchablePlainTexts);
            results.Add("Links", res.Links);
            results.Add("ImageSources", res.ImageSources);

            return results;
        }
        #endregion

        #region Script Editor functions
        /// <summary>
        /// Uses the UseCommunityScriptEditor mapping property provided via the PageTransformationInformation instance to determine the mapping
        /// </summary>
        /// <param name="useQuickLinks">The UseCommunityScriptEditor mapping property provided via the PageTransformationInformation instance</param>
        /// <returns>Whether to transform via the community script editor web part</returns>
        [SelectorDocumentation(Description = "Uses the UseCommunityScriptEditor mapping property provided via the PageTransformationInformation instance to determine the mapping",
                               Example = "ScriptEditorSelector({UseCommunityScriptEditor})")]
        [InputDocumentation(Name = "{UseCommunityScriptEditor}", Description = "The UseCommunityScriptEditor mapping property provided via the PageTransformationInformation instance")]
        [OutputDocumentation(Name = "UseCommunityScriptEditor", Description = "Transform to the community script editor web part")]
        [OutputDocumentation(Name = "NoScriptEditor", Description = "Don't transform as there's no script editor")]
        public string ScriptEditorSelector(string useCommunityScriptEditor)
        {
            if (bool.TryParse(useCommunityScriptEditor, out bool useCommunityScriptEditorBool))
            {
                if (useCommunityScriptEditorBool)
                {
                    return "UseCommunityScriptEditor";
                }
            }

            return "NoScriptEditor";
        }
        #endregion

        #region Contact functions
        /// <summary>
        /// Checks if the passed value is a user or not
        /// </summary>
        /// <param name="person">Account of the user</param>
        /// <returns>Indication if user is valid or not</returns>
        [SelectorDocumentation(Description = "Checks if the passed value is a user or not",
                               Example = "UserExistsSelector({PersonEmail})")]
        [InputDocumentation(Name = "{PersonEmail}", Description = "Account of the user")]
        [OutputDocumentation(Name = "InvalidUser", Description = "User is invalid")]
        [OutputDocumentation(Name = "ValidUser", Description = "User info is valid")]
        public string UserExistsSelector(string person)
        {
            if (string.IsNullOrEmpty(person))
            {
                return "InvalidUser";
            }

            return "ValidUser";
        }

        /// <summary>
        /// Looks up a person from the UserInfo list and returns the needed details
        /// </summary>
        /// <param name="person">User account to lookup (in i:0#.f|membership|joe@contoso.onmicrosoft.com format)</param>
        /// <returns>Information about the found user</returns>
        [FunctionDocumentation(Description = "Looks up a person from the UserInfo list and returns the needed details",
                               Example = "LookupPerson({ContactLoginName})")]
        [InputDocumentation(Name = "{ContactLoginName}", Description = "User account to lookup (in i:0#.f|membership|joe@contoso.onmicrosoft.com format)")]
        [OutputDocumentation(Name = "PersonName", Description = "Name of the user")]
        [OutputDocumentation(Name = "PersonEmail", Description = "User's email")]
        [OutputDocumentation(Name = "PersonUPN", Description = "UPN of the user")]
        [OutputDocumentation(Name = "PersonRole", Description = "Role of the user")]
        [OutputDocumentation(Name = "PersonDepartment", Description = "User's department")]
        [OutputDocumentation(Name = "PersonPhone", Description = "Phone number of the user")]
        [OutputDocumentation(Name = "PersonSip", Description = "SIP address of the user")]
        public Dictionary<string, string> LookupPerson(string person)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            if (string.IsNullOrEmpty(person))
            {
                return result;
            }

            string CAMLQueryByName = @"
                <View Scope='Recursive'>
                  <Query>
                    <Where>
                      <Contains>
                        <FieldRef Name='Name'/>
                        <Value Type='text'>{0}</Value>
                      </Contains>
                    </Where>
                  </Query>
                </View>";

            List siteUserInfoList = this.sourceClientContext.Web.SiteUserInfoList;
            CamlQuery query = new CamlQuery
            {
                ViewXml = String.Format(CAMLQueryByName, person)
            };
            var loadedUsers = this.sourceClientContext.LoadQuery(siteUserInfoList.GetItems(query));
            this.sourceClientContext.ExecuteQueryRetry();

            if (loadedUsers != null)
            {
                var loadedUser = loadedUsers.FirstOrDefault();
                if (loadedUser != null)
                {
                    result.Add("PersonName", loadedUser["Title"] != null ? loadedUser["Title"].ToString() : "");
                    result.Add("PersonEmail", loadedUser["EMail"] != null ? loadedUser["EMail"].ToString() : "");
                    result.Add("PersonUPN", loadedUser["UserName"] != null ? loadedUser["UserName"].ToString() : "");
                    result.Add("PersonRole", loadedUser["JobTitle"] != null ? loadedUser["JobTitle"].ToString() : "");
                    result.Add("PersonDepartment", loadedUser["Department"] != null ? loadedUser["Department"].ToString() : "");
                    result.Add("PersonPhone", loadedUser["WorkPhone"] != null ? loadedUser["WorkPhone"].ToString() : "");
                    result.Add("PersonSip", loadedUser["SipAddress"] != null ? loadedUser["SipAddress"].ToString() : "");
                }
            }
            else
            {
                // Fallback...
                var personParts = person.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                if (personParts.Length == 3)
                {
                    person = personParts[2];
                }

                result.Add("PersonName", "");
                result.Add("PersonEmail", person);
                result.Add("PersonUPN", person);
                result.Add("PersonRole", "");
                result.Add("PersonDepartment", "");
                result.Add("PersonPhone", "");
                result.Add("PersonSip", "");
            }

            return result;
        }
        #endregion

        #region Helper methods
        private bool IsCrossSiteTransfer()
        {
            if (this.sourceClientContext == null)
            {
                return false;
            }

            if (this.sourceClientContext.Web.GetUrl().Equals(this.clientSidePage.Context.Web.GetUrl(), StringComparison.InvariantCultureIgnoreCase))
            {
                return false;
            }

            return true;
        }
        #endregion
    }
}
