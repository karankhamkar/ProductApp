using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Xml.Linq;

namespace ProductApp
{
    public class Product
    {
        #region Properties
        public string Name { get; set; }
        public int BrandId { get; set; }
        public DateTime MfgDate { get; set; }

        private static int productCounter;

        private int productId;
        public int ProductId
        {
            get { return productId; }
        }
        public int ManufactureYear
        {
            get
            {
                return (int)MfgDate.Year;
            }
        }
        public string ManufactureMonth
        {
            get
            {
                return MfgDate.ToString("MMMM");
            }
        }
        public DateTime IsExpire
        {
            get
            {
                return MfgDate.AddMonths(-6);
            }
        }
        public DateTime IsExpiringDate
        {
            get
            {
                return MfgDate.AddMonths(6);
            }
        }
        #endregion

        #region Constructore
        public Product(string name, int brandId, DateTime mfgDate)
        {
            SetValues(name, brandId, mfgDate);
            productId = productCounter;
            productCounter++;
        }
        public Product(int productId, string name, int brandId, DateTime mfgDate)
        {
            SetValues(name, brandId, mfgDate);
            this.productId = productId;
        }
        static Product()
        {
            productCounter = 1;
        }
        #endregion


        #region Methods
        private void SetValues(string name, int brandId, DateTime mfgDate)
        {
            Name = name;
            BrandId = brandId;
            MfgDate = mfgDate;

        }
        public static bool IsNameValid(string input)
        {
            bool result = false;
            if (string.IsNullOrEmpty(input) == false && input.Length <= 100)
            {
                result = true;
            }
            return result;
        }

        //public static bool IsManufactureDateValid(string input, out DateTime mfgDate)
        //{
        //    bool result = false;
        //    mfgDate = mfgDate;
        //    DateTime parseDate;
        //    bool isValidDate = DateTime.TryParse(input, out parseDate);
        //    if(isValidDate)
        //    {
        //        result = true;
        //        mfgDate = parseDate;

        //    }
        //    return result;
        //}

        #endregion

    }
}