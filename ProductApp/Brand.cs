namespace ProductApp
{
    public class Brand
    {
       

        public int Id { get; set; }
        //public string Name { get; set; }
       
        public Brand(int id)
        {
            Id = id;
        }

        public static bool IsValidBrand(string id, out int brandId)
        {
            bool result = false;
            brandId = -1;
            int parsedInteger;
            bool isValidInteger = int.TryParse(id, out parsedInteger);
            if (isValidInteger)
            {
                var brandList = DataStore.GetAllBrands();

                foreach (var item in brandList)
                {
                    if (item.Id == parsedInteger)
                    {
                        result = true;
                        brandId = parsedInteger;
                        break;
                    }
                }
            }

            return result;
        }
    }
}