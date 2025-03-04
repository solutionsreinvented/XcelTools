using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Models.Sections
{
    [ComVisible(true)]
    public class Section
    {
        public string Designation { get; set; }
        public string StaadName { get; set; }
        public double A { get; set; }
        public double Mass { get; set; }
        public double MassFps { get; set; }
        public double ALO { get; set; }
        public double AGO { get; set; }
        public double Cy { get; set; }
        public double Ey { get; set; }
        public double Cz { get; set; }
        public double Ez { get; set; }
        public double Iy { get; set; }
        public double Wely { get; set; }
        public double Wply { get; set; }
        public double Ry { get; set; }
        public double Iz { get; set; }
        public double Welz { get; set; }
        public double Wplz { get; set; }
        public double Rz { get; set; }
        public double Iu { get; set; }
        public double Rv { get; set; }
        public double Iv { get; set; }
        public double It { get; set; }
        public double Iw { get; set; }
    }
}
