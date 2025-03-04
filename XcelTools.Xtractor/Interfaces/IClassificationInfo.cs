using ReInvented.Sections.Domain.Interfaces;

using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Interfaces
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IClassificationInfo
    {
        string Category { get; set; }
        IRolledSection[] Sections { get; set; }
        string TableName { get; set; }
    }
}