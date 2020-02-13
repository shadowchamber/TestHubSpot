using HubSpotLib;
using System;

namespace ConsoleTest
{
    /// <summary>
    /// Main entry point class
    /// </summary>
    class Program
    {
        /// <summary>
        /// entry point
        /// </summary>
        /// <param name="args">console arguments</param>
        static void Main(string[] args)
        {
            try
            {
                Contacts contacts = new Contacts("demo");

                // calling method A (FetchUsers)
                var list = contacts.FetchUsers(new DateTime(2020, 2, 6));

                // calling method B
                contacts.ExportoExcel(list, "test.xlsx");
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
            }
        }
    }
}
