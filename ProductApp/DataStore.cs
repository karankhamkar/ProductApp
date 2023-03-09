using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProductApp
{
    public class DataStore
    {

        #region Methods
        public static List<Brand> GetAllBrands()
        {
            List<Brand> list = new List<Brand>();
            list.Add(new Brand(101));
            list.Add(new Brand(102));
            list.Add(new Brand(103));
            list.Add(new Brand(104));
            return list;
        }
       
        #endregion
    }
}
