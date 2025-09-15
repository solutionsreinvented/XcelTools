using ReInvented.DataAccess.Services;
using ReInvented.StaadPro.Interop.Models;
using ReInvented.StaadPro.Interop.Services;

using System.Runtime.InteropServices;
using XcelTools.Interfaces;
using XcelTools.Staad.Interop.Interfaces;
using XcelTools.Staad.Interop.Models;
using XcelTools.Xtractor.Services;

namespace XcelTools.Models
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class ServicesProvider : IServicesProvider
    {
        public MaterialsService GetMaterialsService()
        {
            return MaterialsService.Instance;
        }

        public SectionsService GetSectionsService()
        {
            return SectionsService.Instance;
        }

        public IOpenStaadWrapperCom GetOpenStaadWrapper()
        {
            string filePath = FileServiceProvider.GetFilePathUsingOpenFileDialog(new ReInvented.DataAccess.Models.FileFilter("Staad Files", FileExtensions.StaadApplication));
            OpenStaadWrapper wrapper = OpenStaadWrapperProvider.Get(filePath);
            return new OpenStaadWrapperCom(wrapper);
        }
    }
}
