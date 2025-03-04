using System.Runtime.InteropServices;

using XcelTools.Xtractor.Models.Sections;

namespace XcelTools.Xtractor.Interfaces
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDatabaseInfo
    {
        string Country { get; set; }
        string Name { get; set; }
        IShapeInfo[] SectionShapes { get; set; }
        object GetSectionShapes();
    }
}