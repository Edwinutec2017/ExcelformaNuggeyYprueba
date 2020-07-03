using System;
using System.Collections.Generic;
using System.Text;

namespace PrueebasNugguet2
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    class DescripcionExcel:Attribute
    {
        public string Name { get; set; }
        public bool Ignore { get; set; }
    }
}
