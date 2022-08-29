using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Models.Sections
{
    [ComVisible(true)]
    public class Database<T> where T : Section, new()
    {
        public string CopyRight { get; set; }

        public List<Table<T>> Tables { get; set; }
    }
}
