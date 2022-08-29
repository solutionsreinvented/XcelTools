using XcelTools.Xtractor.Interfaces;

using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Services
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class MaterialsProvider : IMaterialsProvider
    {
        private static MaterialsProvider _instance;

        private MaterialsProvider()
        {

        }

        public static MaterialsProvider Instance()
        {
            return _instance ?? new MaterialsProvider();
        }
    }
}
