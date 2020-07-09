using System;
using System.Collections.Generic;
using System.Text;

namespace EgenerarExcel2hoja.Dto
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
   public class DescripcionExcel : Attribute
    {
        public string Name { get; set; }
        public bool Ignore { get; set; }
    }
}
