using System.Collections.Generic;
using System.Reflection;
using System.Runtime.Serialization;
using HubSpot.NET.Core.Interfaces;

namespace HubSpot.NET.Api.Company.Dto
{
    /// <summary>
    /// Models a Company entity within HubSpot. Default properties are included here
    /// with the intention that you'd extend this class with properties specific to 
    /// your HubSpot account.
    /// </summary>
    [DataContract]
    public class CompanyHubSpotModel : IHubSpotModel
    {
        [DataMember(Name = "companyId")]
        [IgnoreDataMember]
        public long? Id { get; set; }

        [DataMember(Name = "name")]
        public string Name { get; set; }

        [DataMember(Name = "domain")]
        public string Domain { get; set; }

        [DataMember(Name = "website")]
        public string Website { get; set; }

        [DataMember(Name = "description")]
        public string Description { get; set; }

        [DataMember(Name = "country")]
        public string Country { get; set; }

        [DataMember(Name = "city")]
        public string City { get; set; }

        [DataMember(Name = "state")]
        public string State { get; set; }

        [DataMember(Name = "zip")]
        public string Zip { get; set; }

        [DataMember(Name = "phone")]
        public string Phone { get; set; }

        public bool IsNameValue => true;

        [DataMember(Name = "properties")]
        public Dictionary<string, CompanyProperty> Properties { get; set; } = new Dictionary<string, CompanyProperty>();

        /// <summary>
        /// Loads all properties into the ContactProperty Dictionary even from custom contact models
        /// </summary>
        public void LoadProperties()
        {
            PropertyInfo[] properties = GetType().GetProperties();

            foreach (PropertyInfo prop in properties)
            {
                var key = prop.GetCustomAttribute(typeof(DataMemberAttribute)) as DataMemberAttribute;
                object value = prop.GetValue(this);

                if (value == null || key == null || key.Name == "properties")
                    continue;

                Properties.Add(key.Name, new CompanyProperty(value));
            }
        }

        /// <summary>
        /// Upload all properties from the CompanyProperty Dictionary even to the custom company models
        /// </summary>
        public void UploadProperties()
        {
            PropertyInfo[] properties = GetType().GetProperties();

            foreach (PropertyInfo prop in properties)
            {
                var key = prop.GetCustomAttribute(typeof(DataMemberAttribute)) as DataMemberAttribute;                

                if (key == null || key.Name == "properties")
                    continue;

                if (!Properties.ContainsKey(key.Name))
                {
                    continue;
                }

                CompanyProperty value = Properties[key.Name];

                prop.SetValue(this, value.Value);
            }
        }
    }
}
