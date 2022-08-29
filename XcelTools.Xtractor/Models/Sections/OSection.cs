using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Models.Sections
{
    [ComVisible(true)]
    public class OSection : Section
    {
        public double ID { get; set; }
        public double OD { get; set; }
        public double ALI { get; set; }
        public double AGI { get; set; }
        public double Tw { get; set; }
    }
}
