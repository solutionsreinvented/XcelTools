using ReInvented.Sections.Domain.Interfaces;
using ReInvented.Sections.Domain.Models;

using System.Linq;
using System.Runtime.InteropServices;

using XcelTools.Xtractor.Interfaces;

namespace XcelTools.Xtractor.Models.Sections
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]

    public class ClassificationInfo : IClassificationInfo
    {
        public string Category { get; set; }

        public string TableName { get; set; }

        public IRolledSection[] Sections { get; set; }

        public static ClassificationInfo Transform(Classification c)
        {
            ClassificationInfo classificationInfo = new ClassificationInfo()
            {
                TableName = c.TableName,
                Category = c.Category,
                Sections = c.Sections.ToArray()
            };

            return classificationInfo;
        }
    }
}