using OpenSTAADUI;

using ReInvented.StaadPro.Interop.Extensions;

using System.Runtime.InteropServices;

using XcelTools.Staad.Interop.Interfaces;

namespace XcelTools.Staad.Interop.Models
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class OSGeometryCom : IOSGeometryCom
    {
        #region Parameterized Constructor

        public OSGeometryCom(OSGeometryUI geometry)
        {
            Geometry = geometry;
        }

        #endregion

        #region Private Properties

        private OSGeometryUI Geometry { get; set; }

        #endregion

        #region Public Methods

        [ComVisible(true)]
        public void DeleteAllNodes(int nThreads = 1)
        {
            Geometry.DeleteAllNodes(nThreads);
        }

        #endregion

    }
}
