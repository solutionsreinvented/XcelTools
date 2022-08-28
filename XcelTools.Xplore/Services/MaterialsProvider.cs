using System.Runtime.InteropServices;

namespace XcelTools.Xplore.Services
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class MaterialsProvider
    {
        private static MaterialsProvider _instance;

        private MaterialsProvider()
        {

        }

        public static MaterialsProvider Instance()
        {
            return _instance ?? (_instance = new MaterialsProvider());
        }
    }
}
