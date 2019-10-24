﻿using OfficeDevPnP.Core.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class that deals with site (both site collection and web site) creation, status, retrieval and settings
    /// </summary>
    public static partial class WebExtensions
    {

        /// <summary>
        /// Gets a list of SharePoint lists to scan for modern compatibility
        /// </summary>
        /// <param name="web">Web to check</param>
        /// <returns>List of SharePoint lists to scan</returns>
        public static List<List> GetListsToScan(this Web web, bool showHidden=false)
        {
            List<List> lists = new List<List>(10);

            // Ensure timeout is set on current context as this can be an operation that times out
            web.Context.RequestTimeout = Timeout.Infinite;

            ListCollection listCollection = web.Lists;
            listCollection.EnsureProperties(coll => coll.Include(li=>li.Id, li => li.ForceCheckout, li => li.Title, li => li.Hidden, li => li.DefaultViewUrl, 
                                                                 li => li.BaseTemplate, li => li.RootFolder, li => li.ListExperienceOptions, li => li.ItemCount, 
                                                                 li => li.UserCustomActions, li => li.LastItemUserModifiedDate, li => li.DocumentTemplateUrl, 
                                                                 li => li.IsApplicationList, li => li.IsCatalog, li => li.IsEnterpriseGalleryLibrary, li => li.IsSiteAssetsLibrary,
                                                                 li => li.IsSystemList, li => li.BaseType));

            // Let's process the visible lists
            IQueryable<List> listsToReturn = null;

            if (showHidden)
            {
                listsToReturn = listCollection;
            }
            else
            {
                listsToReturn = listCollection.Where(p => p.Hidden == false);
            }

            foreach (List list in listsToReturn)
            {
                if (list.DefaultViewUrl.Contains("_catalogs"))
                {
                    // skip catalogs
                    continue;
                }

                if (list.BaseTemplate == 544)
                {
                    // skip MicroFeed (544)
                    continue;
                }

                lists.Add(list);
            }

            return lists;
        }    
    }
}
