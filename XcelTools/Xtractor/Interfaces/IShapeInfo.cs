using System.Runtime.InteropServices;

using XcelTools.Xtractor.Enums;
using XcelTools.Xtractor.Interfaces;

namespace XcelTools.Xtractor.Models.Sections
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IShapeInfo
    {
        IClassificationInfo[] Classifications { get; set; }
        string Description { get; set; }
        Shape Shape { get; set; }
    }
}