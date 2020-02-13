using HubSpotLib;
using NUnit.Framework;
using System;

namespace HubSpotLibTests
{
    public class Tests
    {
        Contacts contacts;

        [SetUp]
        public void Setup()
        {
            contacts = new Contacts("demo");
        }

        [Test]
        public void GetCompanyIdFromCompanyName_GetForSomeNameTest()
        {
            long companyId = contacts.GetCompanyIdFromCompanyName("Test");

            Assert.AreEqual(142226976, companyId); // id got from HubSpot demo data
        }

        [Test]
        public void FetchUsers_SimpleTest()
        {
            var list = contacts.FetchUsers(new DateTime(2000, 1, 1));

            Assert.AreEqual(13, list.Count); 
        }
    }
}