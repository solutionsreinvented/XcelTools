using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Models.Sections
{
    [ComVisible(true)]
    public class Table<T> where T : Section, new()
    {
        public string Name { get; set; }

        public List<T> Profiles { get; set; }
    }
}
