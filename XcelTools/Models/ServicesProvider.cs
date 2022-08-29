using System.Runtime.InteropServices;
using XcelTools.Interfaces;
using XcelTools.Xtractor.Services;

namespace XcelTools.Models
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class ServicesProvider : IServicesProvider
    {
        public MaterialsProvider GetMaterialsProvider()
        {
            return MaterialsProvider.Instance();
        }

        public SectionsProvider GetSectionsProvider()
        {
            return SectionsProvider.Instance();
        }
    }
}
