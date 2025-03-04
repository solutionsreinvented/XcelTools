using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Models.Materials
{
    [ComVisible(true)]
    public class CommonProperties
    {
        public double E { get; set; }

        public double Poisson { get; set; }

        public double Alpha { get; set; }

        public double Density { get; set; }

        public double CrDamp { get; set; }
    }
}