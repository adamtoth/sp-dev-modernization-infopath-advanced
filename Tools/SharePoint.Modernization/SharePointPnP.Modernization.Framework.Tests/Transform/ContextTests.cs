﻿using System;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SharePointPnP.Modernization.Framework.Tests.Transform
{
    [TestClass]
    public class ContextTests
    {
        [TestMethod]
        public void ContextSPO_BasicTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var web = clientContext.Web;
                var title = web.EnsureProperty(o => o.Title);
                Console.WriteLine(title);
                Assert.IsTrue(!string.IsNullOrEmpty(title)); 
            }
        }

        [TestMethod]
        public void ContextSPOnPremises_BasicTest()
        {
            using (var clientContext = TestCommon.CreateOnPremisesClientContext())
            {
                var web = clientContext.Web;
                var title = web.EnsureProperty(o => o.Title);
                Console.WriteLine(title);
                Assert.IsTrue(!string.IsNullOrEmpty(title));
            }
        }
    }
}
