using XcelTools.Xtractor.Models.Sections;

using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace XcelTools.Xtractor.Interfaces
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]

    public interface ISectionsService
    {
        [ComVisible(true)]
        void PopulateSectionClassifications(string dbName, string shape, Excel.Range rngTarget);

        [ComVisible(true)]
        void PopulateSectionProfiles(string dbName, string shape, string classification, Excel.Range rngTarget, Excel.Range rngPopulation);

        [ComVisible(true)]
        IClassificationInfo GetClassification(ShapeInfo shapeInfo, string category);

        [ComVisible(true)]
        IShapeInfo GetShapeInfo(DatabaseInfo databaseInfo, string shape);

        [ComVisible(true)]
        IDatabaseInfo GetDatabaseInfo(string country);

        [ComVisible(true)]
        Catalogue GetSectionsCatalogue();
    }
}
