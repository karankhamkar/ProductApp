using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace Demo
{
    public class Product
    {
        #region Properties
        public string ProductName { get; set; }
        public double Price { get; set; }
        public string BrandName { get; set; }
        public DateTime ManufacturingDate { get; set; }

        #endregion

        #region Methods
        public static bool IsNameValid(string input)
        {
            bool result = false;
            if (string.IsNullOrEmpty(input) == false && input.Length <= 50)
            {
                result = true;
            }
            return result;
        }

        public static bool IsBrandNameValid(string input)
        {
            bool result = false;
            if (string.IsNullOrEmpty(input) == false && input.Length <= 50)
            {
                result = true;
            }
            return result;
        }

        public static bool IsManufactureDateValid(string mfgDate)
        {
            DateTime parseDate;
            bool isCorrectDate = DateTime.TryParse(mfgDate, out parseDate);
            if (isCorrectDate)
            {
                return true;
            }
            return false;

        }

        #endregion


    }
}
