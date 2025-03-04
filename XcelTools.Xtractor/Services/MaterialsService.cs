using XcelTools.Xtractor.Interfaces;

using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Services
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class MaterialsService : IMaterialsProvider
    {
        private static MaterialsService _instance;

        private MaterialsService()
        {

        }

        public static MaterialsService Instance => _instance ?? (_instance = new MaterialsService());

    }
}
