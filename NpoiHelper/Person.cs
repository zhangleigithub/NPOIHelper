using NPOIEx;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NpoiHelper
{
    public class Person
    {
        [ExcelColumn(0)]
        public string Id { get; set; }

        [ExcelColumn(1)]
        public string Name { get; set; }
    }
}
