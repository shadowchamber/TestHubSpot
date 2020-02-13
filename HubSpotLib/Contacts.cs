using ClosedXML.Excel;
using HubSpot.NET.Api.Company.Dto;
using HubSpot.NET.Api.Contact.Dto;
using HubSpot.NET.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HubSpotLib
{
    /// <summary>
    /// Class for additional HubSpot contacts operations
    /// </summary>
    public class Contacts
    {
        /// <summary>
        /// Initializes contacts instance
        /// </summary>
        /// <param name="apikey">api key value</param>
        public Contacts(string apikey)
        {
            apiKey = apikey;
        }

        private string apiKey;

        private HubSpotApi api;
        private HubSpotApi API
        {
            get
            {
                // lazy init
                if (api == null)
                {
                    api = new HubSpotApi(apiKey);
                }

                return api;
            }
        }

        /// <summary>
        /// companies list
        /// </summary>
        private CompanyListHubSpotModel<CompanyHubSpotModel> companies = new CompanyListHubSpotModel<CompanyHubSpotModel>();
        
        /// <summary>
        /// shows if all companies loaded to companies list
        /// </summary>
        private bool allCompaniesLoaded = false;

        /// <summary>
        /// returns company id by company name
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public long GetCompanyIdFromCompanyName(string name)
        {
            if (companies.Companies == null)
            {
                companies.Companies = new List<CompanyHubSpotModel>();
            }

            // find if company in list
            if (companies.Companies.Where(i => i.Name == name).Count() > 0)
            {
                var companyFound = companies.Companies.Where(i => i.Name == name).First();

                return Convert.ToInt64(companyFound.Id);
            }

            // if not we'll take it from hubspot
            while (!allCompaniesLoaded)
            {                 
                // 100 is max count
                var companiesResponse = API.Company.List(new ListRequestOptions() { Limit = 100, Offset = companies.ContinuationOffset });

                // if hubspot doesn't return more companies it means we have full list
                if (companiesResponse.Companies.Count == 0)
                {
                    allCompaniesLoaded = true;
                }

                // take companies to full list
                foreach (var company in companiesResponse.Companies)
                {
                    // if name didn't came
                    if (company.Name ==  null || company.Name == "")
                    {
                        // we load it from property list
                        company.UploadProperties();

                        // if it is still not loaded
                        if (company.Name == null || company.Name == "")
                        {
                            // we take all company data from hubspot
                            var companyUpdated = api.Company.GetById(Convert.ToInt64(company.Id));

                            // loaded properties to fields
                            companyUpdated.UploadProperties();

                            companies.Companies.Add(companyUpdated);

                            // store offset for next part of the list
                            companies.ContinuationOffset = Convert.ToInt64(company.Id);

                            continue;
                        }
                    }

                    // if we have name in field
                    companies.Companies.Add(company);

                    // store offset for next part of the list
                    companies.ContinuationOffset = Convert.ToInt64(company.Id);
                }

                if (companies.Companies.Where(i => i.Name == name).Count() > 0)
                {
                    var companyFound = companies.Companies.Where(i => i.Name == name).First();

                    return Convert.ToInt64(companyFound.Id);
                }
            }

            return 0;
        }

        private void SetCellStyle(IXLCell cell)
        {
            cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
        }

        public void ExportoExcel(List<Contact> contacts, string filename)
        {
            // we use closed xml cause it is more flexible and faster than for example COM
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Report");

                int i = 1;

                // reuirements for formatting weren't specified so I put only simple steps to show how it could be done
                worksheet.Column("A").Width = 20;
                worksheet.Column("B").Width = 20;
                worksheet.Column("C").Width = 20;
                worksheet.Column("D").Width = 20;
                worksheet.Column("E").Width = 20;
                worksheet.Column("F").Width = 20;
                worksheet.Column("G").Width = 20;
                worksheet.Column("H").Width = 20;
                worksheet.Column("I").Width = 20;
                worksheet.Column("J").Width = 20;
                worksheet.Column("K").Width = 20;

                SetCellStyle(worksheet.Cell("A" + i));
                SetCellStyle(worksheet.Cell("B" + i));
                SetCellStyle(worksheet.Cell("C" + i));
                SetCellStyle(worksheet.Cell("D" + i));
                SetCellStyle(worksheet.Cell("E" + i));
                SetCellStyle(worksheet.Cell("F" + i));
                SetCellStyle(worksheet.Cell("G" + i));
                SetCellStyle(worksheet.Cell("H" + i));
                SetCellStyle(worksheet.Cell("I" + i));
                SetCellStyle(worksheet.Cell("J" + i));
                SetCellStyle(worksheet.Cell("K" + i));

                worksheet.Cell("A" + i).Value = "vid";
                worksheet.Cell("B" + i).Value = "firstname";
                worksheet.Cell("C" + i).Value = "lastname";
                worksheet.Cell("D" + i).Value = "lifecyclestage";
                worksheet.Cell("E" + i).Value = "company_id";
                worksheet.Cell("F" + i).Value = "name";
                worksheet.Cell("G" + i).Value = "website";
                worksheet.Cell("H" + i).Value = "city";
                worksheet.Cell("I" + i).Value = "state";
                worksheet.Cell("J" + i).Value = "zip";
                worksheet.Cell("K" + i).Value = "phone";

                i++;

                foreach (var contact in contacts)
                {
                    worksheet.Cell("A" + i).Value = contact.vid;
                    worksheet.Cell("B" + i).Value = contact.firstname;
                    worksheet.Cell("C" + i).Value = contact.lastname;
                    worksheet.Cell("D" + i).Value = contact.lifecyclestage;
                    worksheet.Cell("E" + i).Value = contact.company_id;
                    worksheet.Cell("F" + i).Value = contact.name;
                    worksheet.Cell("G" + i).Value = contact.website;
                    worksheet.Cell("H" + i).Value = contact.city;
                    worksheet.Cell("I" + i).Value = contact.state;
                    worksheet.Cell("J" + i).Value = contact.zip;
                    worksheet.Cell("K" + i).Value = contact.phone;

                    SetCellStyle(worksheet.Cell("A" + i));
                    SetCellStyle(worksheet.Cell("B" + i));
                    SetCellStyle(worksheet.Cell("C" + i));
                    SetCellStyle(worksheet.Cell("D" + i));
                    SetCellStyle(worksheet.Cell("E" + i));
                    SetCellStyle(worksheet.Cell("F" + i));
                    SetCellStyle(worksheet.Cell("G" + i));
                    SetCellStyle(worksheet.Cell("H" + i));
                    SetCellStyle(worksheet.Cell("I" + i));
                    SetCellStyle(worksheet.Cell("J" + i));
                    SetCellStyle(worksheet.Cell("K" + i));

                    i++;
                }

                workbook.SaveAs(filename);
            }

            // Open Excel using:
            // Microsoft.Office.Interop.Excel
            // Application excel = new Application();
            // Workbook wb = excel.Workbooks.Open(filename);
        }

        /// <summary>
        /// converts unix time to DateTime object
        /// </summary>
        /// <param name="unixTimeStamp">unix time (milliseconds)</param>
        /// <returns>date time object</returns>
        public DateTime UnixTimeStampToDateTime(double unixTimeStamp)
        {
            // Unix timestamp is seconds past epoch
            System.DateTime dtDateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc);
            dtDateTime = dtDateTime.AddMilliseconds(unixTimeStamp).ToLocalTime();

            return dtDateTime;
        }

        /// <summary>
        /// Takes contacts modified after specified date
        /// </summary>
        /// <param name="modifiedOnOrAfter">date after which we should take contacts</param>
        /// <returns>contact list</returns>
        public List<Contact> FetchUsers(DateTime modifiedOnOrAfter)
        {
            List<Contact> res = new List<Contact>();
             
            // take first part of the list
            ContactListHubSpotModel<ContactHubSpotModel> cresponse = API.Contact.RecentlyUpdated(new ListRecentRequestOptions() { Limit = 100, Offset = 0 });

            while (cresponse.Contacts.Count() > 0)
            {
                foreach (var contact in cresponse.Contacts)
                {
                    // store offset for next list part
                    cresponse.ContinuationOffset = Convert.ToInt64(contact.Id);

                    // fail if we don't have last modified date
                    if (!contact.Properties.ContainsKey("lastmodifieddate"))
                    {
                        throw new Exception("last modified date is not specified");
                    }

                    // convert unix tome to datetime
                    string lastModStr = (string)contact.Properties["lastmodifieddate"].Value;
                    DateTime lastMod = UnixTimeStampToDateTime(Convert.ToInt64(lastModStr));

                    // check if date is not in range
                    // we sttop the list cause list is desc sorted so next dates will not be in range as well
                    if (lastMod < modifiedOnOrAfter)
                    {
                        return res; // we don't need the rest
                    }

                    Contact newContact = new Contact();

                    res.Add(newContact);

                    newContact.vid = Convert.ToString(contact.Id);

                    // takes required properties to fields
                    if (contact.Properties.ContainsKey("firstname"))
                    {
                        newContact.firstname = (string)contact.Properties["firstname"].Value;
                    }

                    if (contact.Properties.ContainsKey("lastname"))
                    {
                        newContact.lastname = (string)contact.Properties["lastname"].Value;
                    }

                    if (contact.Properties.ContainsKey("lifecyclestage"))
                    {
                        newContact.lifecyclestage = (string)contact.Properties["lifecyclestage"].Value;
                    }

                    // if company id and company name are not specified
                    if (contact.AssociatedCompanyId == null
                        && !contact.Properties.ContainsKey("company"))
                    {
                        continue;
                    }

                    // if we have name but don't have company id
                    if (contact.AssociatedCompanyId == null
                        && contact.Properties.ContainsKey("company"))
                    {
                        long cmpId = GetCompanyIdFromCompanyName((string)contact.Properties["company"].Value);

                        if (cmpId != 0)
                        {
                            contact.AssociatedCompanyId = cmpId;
                        }
                    }

                    if (contact.AssociatedCompanyId != null)
                    {
                        try
                        {
                            // could fail if company with specified id doen't exist so we use try catch
                            var company = api.Company.GetById(Convert.ToInt64(contact.AssociatedCompanyId));

                            // takes properties to fields
                            company.UploadProperties();

                            if (company != null)
                            {
                                newContact.company_id = Convert.ToString(company.Id);
                                newContact.name = company.Name;
                                newContact.website = company.Website;
                                newContact.city = company.City;
                                newContact.state = company.State;
                                newContact.zip = company.Zip;
                                newContact.phone = company.Phone;
                            }
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine(exception.Message);
                        }
                    }
                }

                // get next list part
                cresponse = API.Contact.RecentlyUpdated(new ListRecentRequestOptions() { Limit = 100, Offset = cresponse.ContinuationOffset });
            }

            return res;
        }
    }
}
