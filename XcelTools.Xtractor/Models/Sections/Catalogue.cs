using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Models.Sections
{
    [ComVisible(true)]
    public class Catalogue
    {
        public string CopyRight { get; set; }

        public string FileExtension { get; set; }

        public List<DataFile> DataFiles { get; set; }

        public List<DatabaseInfo> Databases { get; set; }
    }
}
