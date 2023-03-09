using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Demo
{
    public class Program
    {

        static void Main(string[] args)
        {
            Console.WriteLine("---------------------------- Welcome to Product Store -----------------------------");
            List<Product> products = new List<Product>();
            //Dictionary<string, Product> products = new Dictionary<string, Product>();
           
            char repete;
            do
            {
                Product product= new Product();

                product.ProductName = GetName();
                product.BrandName = GetBrandName();
                product.Price = GetPrice();
                product.ManufacturingDate = GetManufacturingDate();

                products.Add(product);
                Console.WriteLine("Do you Want to Continue, press y / n : ");
                repete = char.Parse(Console.ReadLine());
            }
            while (repete == 'y' || repete == 'Y');

            List<Mobile> mobiles = new List<Mobile>();
           // Dictionary<string, Mobile> mobiles = new Dictionary<string, Mobile>();

            mobiles = GetMobile(products);

            DisplayMobiles();
            
            Console.ReadLine();

            void DisplayMobiles()
            {
                foreach (Mobile item in mobiles)
                {
                    if (item.MobileName.Length > 10)
                    {
                        item.MobileName = item.MobileName.Substring(0, 7) + "...";
                    }

                    if (item.MobileBrand.Length > 10)
                    {
                        item.MobileBrand = item.MobileBrand.Substring(0, 7) + "...";
                    }

                    Console.WriteLine($" Mobile Name : {item.MobileName}\n Mobile Price : {item.MobilePrice}\n Mobile Brand : {item.MobileBrand}\n Mobile Manufacturing Date : {item.MobileManufacturingDate}\n Mobile Is Expiring On : {item.IsExpiringOn}");
                    if (item.IsExpire)
                    {
                        Console.WriteLine("Mobile Manufacture date is Expire.");
                    }
                    else
                    {
                        Console.WriteLine("Mobile Manufacture date Expire date dose not future");
                    }

                    if (item.IsExpensive)
                    {
                        Console.WriteLine("Mobile is Expensive.");
                    }
                    else
                    {
                        Console.WriteLine("Mobile is not Expensive.");
                    }
                }
            }
        }

       

        #region Methods
        private static string GetName()
        {
            
            Console.WriteLine("Enter Product Name : ");
            string input = Console.ReadLine();
            if (Product.IsNameValid(input))
            {
                return input;
            }
            else
            {
                Console.WriteLine("Product name exceeds the character limit.\nPlease Re-enter name.");
                return GetName();
            }
        }
        private static string GetBrandName()
        {
            Console.WriteLine("Enter Product Brand : ");
            string input = Console.ReadLine();
            if(Product.IsBrandNameValid(input))
            {
                return input;
            }
            else
            {
                Console.WriteLine("Invalid input.\nPlease Re-enter Brand Name.");
                return GetBrandName();
            }
        }
        private static double GetPrice()
        {
            Console.WriteLine("Enter Product Price : ");
            double price;
            bool isPrice = double.TryParse(Console.ReadLine(), out price);
            if(isPrice)
            {
                return price;
            }
            else
            {
                Console.WriteLine("Invalid Input.\nPlease Re-enter Price.");
                return GetPrice();
            }
        }
        private static DateTime GetManufacturingDate()
        {
            Console.WriteLine("Enter Manufacture Date(MM-DD-YYYY) : ");
            string mfgDate = Console.ReadLine();
            if (Product.IsManufactureDateValid(mfgDate))
            {
                return DateTime.Parse(mfgDate);
            }
            else
            {
                Console.WriteLine("Invalid Date.\nPlease Re-enter Date.");

            }
            return GetManufacturingDate();

        }

        static Product SearchProduct(List<Product> allProducts, string name)
        {
            foreach(Product item in allProducts)
            {
                if(item.ProductName == name)
                {
                    return item;
                }
            }
            return null;
        }
        //private static Dictionary<string, Mobile> GetMobile(Dictionary<string, Product> products)
        //{
        //    Dictionary<string, Mobile> mobileList = new Dictionary<string, Mobile>();
        //    foreach (KeyValuePair<string, Product> product in products)
        //    {
        //        Mobile mobile = new Mobile();

        //        mobile.MobileName = product.Value.ProductName;
        //        mobile.MobileBrand = product.Value.BrandName;
        //        mobile.MobilePrice = product.Value.Price;
        //        mobile.MobileManufacturingDate = product.Value.ManufacturingDate;

        //        mobileList.Add(mobile.MobileName, mobile);
        //    }
        //    return mobileList;
        //}

        private static List<Mobile> GetMobile(List<Product> products)
        {
            List<Mobile> mobileList = new List<Mobile>(); 

            foreach (Product product in products)
            {
                Mobile mobile = new Mobile();

                mobile.MobileName = product.ProductName;
                mobile.MobileBrand = product.BrandName;
                mobile.MobilePrice = product.Price;
                mobile.MobileManufacturingDate = product.ManufacturingDate;

                mobileList.Add(mobile);   
            }
            return mobileList;
        }
        //private static Mobile GetMobile(Product product)
        //{
        //    Mobile mobile = new Mobile();
        //    mobile.MobileName = product.ProductName;
        //    mobile.MobilePrice = product.Price;
        //    mobile.MobileBrand = product.BrandName;
        //    return mobile;
        //}
        #endregion

    }
}
