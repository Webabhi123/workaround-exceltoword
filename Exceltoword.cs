using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.Model
{
    class Exceltoword
    {
        public class ExcelRowData
        {
            public string Name { get; set; }
            public string  Contact { get; set; }
            public string District { get; set; }
            public string Village { get; set; }
            public string DateofInstallation { get; set; }
            public string StoveSerialNumber { get; set; }
            public string TotalHouseHoldMember { get; set; }
            public string Adults { get; set; }
            public string NumberofChildren { get; set; }

            public string Date { get; set; }

            public string CookstoveModel { get; set; }

            public string pricepaidofcookstove { get; set; }
            //public DateTime DateTime { get; set; }
            //public string Date // Add a new property to get the date without time
            // {
            //   get { return DateTime.ToString("yyyy-MM-dddd"); }
            //}

            // Add more properties for other columns as needed
            // Example: public string Column2 { get; set; }
        }
    }
}
