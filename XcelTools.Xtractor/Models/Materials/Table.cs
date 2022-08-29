using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Models.Materials
{
    [ComVisible(true)]
    public class Table
    {
        public string Name { get; set; }

        public string Country { get; set; }

        public List<Grade> Grades { get; set; }
    }
}
