﻿using System;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SharePointPnP.Modernization.Framework.Tests.Transform.OnPremises
{
    [TestClass]
    public class OnPremisesUtilityTests
    {
        [TestMethod]
        public void Basic_GetListByNameTest()
        {
            TestGetListByName("Pages");
        }

        [TestMethod]
        public void Basic_GetListByName_UpperCaseTest()
        {
            TestGetListByName("PAGES");
        }

        [TestMethod]
        public void Basic_GetListByName_LowerCaseTest()
        {
            TestGetListByName("pages");
        }

        private void TestGetListByName(string listName)
        {
            using (var context = TestCommon.CreateOnPremisesClientContext(TestCommon.AppSetting("SPOnPremDevSiteUrl")))
            {
                var list = context.Web.GetListByName(listName);
                Assert.IsNotNull(list);
            }
        }
    }
}
