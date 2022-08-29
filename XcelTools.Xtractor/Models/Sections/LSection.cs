using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Models.Sections
{
    [ComVisible(true)]
    public class LSection : Section
    {
        public double L { get; set; }
        public double B { get; set; }
        public double R1 { get; set; }
        public double R2 { get; set; }
        public double Tw { get; set; }
    }
}
