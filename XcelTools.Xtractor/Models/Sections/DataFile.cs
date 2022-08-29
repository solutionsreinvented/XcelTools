using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Models.Sections
{
    [ComVisible(true)]
    public class DataFile
    {
        public string Shape { get; set; }

        public string FileName { get; set; }

        public string Description { get; set; }
    }
}
