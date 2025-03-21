using ReInvented.Sections.Domain.Models;

using System;
using System.Runtime.InteropServices;

using XcelTools.Xtractor.Interfaces;

namespace XcelTools.Xtractor.Models.Sections
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]

    public class DatabaseInfo : IDatabaseInfo
    {
        public string Name { get; set; }

        public string Country { get; set; }

        [ComVisible(true)]
        public IShapeInfo[] SectionShapes { get; set; }

        public object GetSectionShapes()
        {
            return SectionShapes;
        }

        public static IDatabaseInfo Transform(Database database)
        {
            DatabaseInfo dbInfo = new DatabaseInfo()
            {
                Country = database.Country,
                Name = database.Name,
                SectionShapes = new ShapeInfo[database.SectionShapes.Count]
            };

            for (int i = 0; i < database.SectionShapes.Count; i++)
            {
                dbInfo.SectionShapes[i] = ShapeInfo.Transform(database.SectionShapes[i]);
            }

            return dbInfo;
        }
    }
}
