using System.Runtime.InteropServices;
using XcelTools.Xplore.Services;

namespace XcelTools.Interfaces
{
    [ComVisible(true)]
    public interface IServicesProvider
    {
        [ComVisible(true)]
        SectionsProvider GetSectionsProvider();

        [ComVisible(true)]
        MaterialsProvider GetMaterialsProvider();
    }
}
