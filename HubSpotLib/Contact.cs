using System;
using System.Collections.Generic;
using System.Text;

namespace HubSpotLib
{
    /// <summary>
    /// Contact with company info
    /// </summary>
    public class Contact
    {
        /// <summary>
        /// Contact identificator
        /// </summary>
        public string vid { get; set; }
        /// <summary>
        /// Firs name
        /// </summary>
        public string firstname { get; set; }
        /// <summary>
        /// Last name
        /// </summary>
        public string lastname { get; set; }
        /// <summary>
        /// Life cycle stage
        /// </summary>
        public string lifecyclestage { get; set; }
        /// <summary>
        /// Company identificatoor
        /// </summary>
        public string company_id { get; set; }
        /// <summary>
        /// Name
        /// </summary>
        public string name { get; set; }
        /// <summary>
        /// Website
        /// </summary>
        public string website { get; set; }
        /// <summary>
        /// City
        /// </summary>
        public string city { get; set; }
        /// <summary>
        /// State
        /// </summary>
        public string state { get; set; }
        /// <summary>
        /// Zip
        /// </summary>
        public string zip { get; set; }
        /// <summary>
        /// Phone
        /// </summary>
        public string phone { get; set; }
    }
}
