using XcelTools.Xtractor.Enums;

using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Models.Sections
{
    [ComVisible(true)]
    public class ShapeInfo
    {
        public Shape Shape { get; set; }

        public string Description { get; set; }

        public List<Classification> Classifications { get; set; }
    }
}
