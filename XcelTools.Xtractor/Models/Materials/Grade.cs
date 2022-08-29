using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Models.Materials
{
    [ComVisible(true)]
    public class Grade
    {
        public string Designation { get; set; }

        public double Fu { get; set; }

        public double Fy { get; set; }

        public string StaadName { get; set; }
    }
}