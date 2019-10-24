using SharePoint.Scanning.Framework;
using System;

namespace SharePoint.Modernization.Scanner.Results
{
    public class InfoPathScanResult: Scan
    {
        public string ListUrl { get; set; }

        public string ListTitle { get; set; }

        public Guid ListId { get; set; }

        /// <summary>
        ///  Indicates how InfoPath is used here: form library or customization of the list form pages
        /// </summary>
        public string InfoPathUsage { get; set; }

        public string InfoPathTemplate { get; set; }

        public bool Enabled { get; set; }

        public int ItemCount { get; set; }

        public DateTime LastItemUserModifiedDate { get; set; }

        public string Mode { get; set; }

        public string ProductVersion { get; set; }

        public string InfoPathTemplateUrl { get; set; }

        public string ContentTypeName { get; set; }

        public string DownloadedXsnId { get; set; }

        public bool HasPersonField { get; set; }

        public bool HasExternalField { get; set; }

        public bool HasSOAPConnection { get; set; }

        public bool HasRESTConnection { get; set; }

        public bool HasDBConnection { get; set; }

        public bool HasRepeatingTable { get; set; }

        public bool HasRepeatingSection { get; set; }

        public bool HasRepeatingRecursiveSection { get; set; }

        public bool HasChoiceGroup { get; set; }

        public bool HasOptionalSection { get; set; }

        public bool HasMasterDetail { get; set; }

        public bool HasRepeatingChoiceGroup { get; set; }

        public bool HasChoiceSection { get; set; }

        public bool HasHorizontalRepeatingTable { get; set; }

        public bool HasDigitalSignature { get; set; }

        public bool HasCodeBehind { get; set; }

        public bool HasPageBreak { get; set; }

        public bool HasMultipleViews { get; set; }

        public bool HasInk { get; set; }

    }
}
