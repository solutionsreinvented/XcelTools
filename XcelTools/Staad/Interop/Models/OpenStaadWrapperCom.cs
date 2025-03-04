using ReInvented.StaadPro.Interop.Models;

using System.Runtime.InteropServices;

using XcelTools.Staad.Interop.Interfaces;

namespace XcelTools.Staad.Interop.Models
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class OpenStaadWrapperCom : IOpenStaadWrapperCom
    {
        #region Parameterized Constructor

        public OpenStaadWrapperCom(OpenStaadWrapper wrapper)
        {
            Wrapper = wrapper;
            Geometry = new OSGeometryCom(wrapper.Geometry);
        }

        #endregion

        #region Private Properties

        private OpenStaadWrapper Wrapper { get; set; }

        #endregion

        #region Public Properties

        public IOSGeometryCom Geometry { get; private set; }

        #endregion
    }
}
