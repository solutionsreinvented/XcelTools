using System.Runtime.InteropServices;

namespace XcelTools.Staad.Interop.Interfaces
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IOSGeometryCom
    {
        [ComVisible(true)]
        void DeleteAllNodes(int nThreads = 1);
    }
}
