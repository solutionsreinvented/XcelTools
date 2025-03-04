
using System.Runtime.InteropServices;

using XcelTools.Staad.Interop.Interfaces;
using XcelTools.Xtractor.Services;

namespace XcelTools.Interfaces
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IServicesProvider
    {
        [ComVisible(true)]
        SectionsService GetSectionsService();

        [ComVisible(true)]
        MaterialsService GetMaterialsService();

        [ComVisible(true)]
        IOpenStaadWrapperCom GetOpenStaadWrapper();
    }
}
