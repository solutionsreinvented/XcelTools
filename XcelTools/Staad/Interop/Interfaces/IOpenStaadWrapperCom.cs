using System.Runtime.InteropServices;

namespace XcelTools.Staad.Interop.Interfaces
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IOpenStaadWrapperCom
    {
        [ComVisible(true)]
        IOSGeometryCom Geometry { get; }
    }
}
