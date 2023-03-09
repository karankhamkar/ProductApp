using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Demo
{
    public class Mobile
    {
        #region Properties
        private string _mobileName;

        public string MobileName
        {
            get { return _mobileName; }
            set { _mobileName = value; }
        }

        private Double _mobilePrice;

        public Double MobilePrice
        {
            get { return _mobilePrice; }
            set { _mobilePrice = value; }
        }

        private string _mobileBrand;

        public string MobileBrand
        {
            get { return _mobileBrand; }
            set { _mobileBrand = value; }
        }

        private DateTime _mobileManufacturingDate;

        public DateTime MobileManufacturingDate
        {
            get { return _mobileManufacturingDate; }
            set { _mobileManufacturingDate = value; }
        }
      
        public bool IsExpire
        {
            get
            {
                if(DateTime.Today > MobileManufacturingDate.AddMonths(6))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        public DateTime IsExpiringOn
        {
            get
            {
                return MobileManufacturingDate.AddMonths(6);
            }
        }

        public bool IsExpensive
        {
            get
            {
                if(MobilePrice > 100000)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }
        #endregion
    }
}
