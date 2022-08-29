using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Models.Sections
{
    [ComVisible(true)]
    public class HCSection : Section
    {
        public double H { get; set; }
        public double Bf { get; set; }
        public double R1 { get; set; }
        public double R2 { get; set; }
        public double Tf { get; set; }
        public double Tw { get; set; }
        public double Alpha { get; set; }
    }
}
