using XcelTools.Xtractor.Enums;
using System.Runtime.InteropServices;
using ReInvented.Sections.Domain.Models;
using System;
using XcelTools.Xtractor.Interfaces;

namespace XcelTools.Xtractor.Models.Sections
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]

    public class ShapeInfo : IShapeInfo
    {
        //public string Shape { get; set; }
        public Shape Shape { get; set; }

        public string Description { get; set; }

        public IClassificationInfo[] Classifications { get; set; }

        public static IShapeInfo Transform(SectionShape s)
        {
            ShapeInfo shapeInfo = new ShapeInfo()
            {
                Shape = (Shape)Enum.Parse(typeof(Shape), s.Shape),
                Description = s.Description,
                Classifications = new ClassificationInfo[s.Classifications.Count]
            };

            for (int i = 0; i < s.Classifications.Count; i++)
            {
                shapeInfo.Classifications[i] = ClassificationInfo.Transform(s.Classifications[i]);
            }

            return shapeInfo;
        }
    }
}
