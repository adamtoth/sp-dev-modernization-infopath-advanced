﻿namespace SharePointPnP.Modernization.Framework.Publishing.Layouts
{
    /// <summary>
    /// Class for holding data properties for the fields that will be used in the page header
    /// </summary>
    internal class PageLayoutHeaderFieldEntity
    {
        internal string Type { get; set; }
        internal string Name { get; set; }
        internal string HeaderProperty { get; set; }
        internal string Functions { get; set; }
        internal string Alignment { get; set; }
        internal bool ShowPublishedDate { get; set; }
    }
}
