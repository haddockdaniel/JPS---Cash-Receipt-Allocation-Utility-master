using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Windows.Forms;
using JurisAuthenticator;
using System.ComponentModel;
using System.Threading;

namespace JurisUtilityBase
{
    public class DataError
    {
        public DataError()
        {
            rowNum = 0;
            error = "";
        }

        public int rowNum { get; set; }

        public string error { get; set; }

       


    }


}
