using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelNugget02
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
   public  class DescripcionExcel:Attribute
    {
        public string Name { get; set; }
        public bool Ignore { get; set; }
    }
}
