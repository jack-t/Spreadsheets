using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Services
{
    public class Fee
    {
        public Fee(string state, string county, string product, int price, bool pending)
        {
            State = state;
            County = county;
            ProductTypeName = product;
            CurrentPrice = price;
            Pending = pending;
        }

        public string State { get; }
        public string County { get; }
        public string ProductTypeName { get; }
        public int CurrentPrice { get; }
        public bool Pending { get; }
    }
}
