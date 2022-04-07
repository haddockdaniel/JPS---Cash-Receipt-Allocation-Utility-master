using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JurisUtilityBase
{
    public class FeeExpAlloc
    {
        public FeeExpAlloc()
        {
            billNo = 0;
            mat = 0;
            tkpr = 0;
            code = "";
            act = "";
            amt = 0.00;
            allocAmt = 0.00;
            pct = 0.00;
        }

        public int billNo { get; set; }
        public int mat { get; set; }
        public int tkpr { get; set; }
        public string code { get; set; }

        public string act { get; set; }

        public double amt { get; set; }

        public double allocAmt { get; set; }

        public double pct { get; set; }
    }
}
