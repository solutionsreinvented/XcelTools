using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Models.Sections
{
    [ComVisible(true)]
    public class DatabaseInfo
    {
        public string Name { get; set; }

        public string Country { get; set; }

        public List<ShapeInfo> SectionShapes { get; set; }
    }
}
