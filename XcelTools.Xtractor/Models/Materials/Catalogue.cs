using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Models.Materials
{
    [ComVisible(true)]
    public class Catalogue
    {
        public string CopyRight { get; set; }

        public CommonProperties CommonProperties { get; set; }

        public List<Table> Tables { get; set; }
    }
}
