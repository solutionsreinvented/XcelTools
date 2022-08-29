using XcelTools.Xtractor.Models.Sections;

using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Interfaces
{
    [ComVisible(true)]
    public interface ISectionsProvider
    {
        [ComVisible(true)]
        Classification GetClassification(ShapeInfo shapeInfo, string category);

        [ComVisible(true)]
        ShapeInfo GetShapeInfo(DatabaseInfo databaseInfo, string shape);

        [ComVisible(true)]
        DatabaseInfo GetDatabaseInfo(Catalogue catalogue, string country);

        [ComVisible(true)]
        Catalogue GetSectionsCatalogue();

        [ComVisible(true)]
        Database<T> GetDatabase<T>(string fileName) where T : Section, new();

        [ComVisible(true)]
        Table<T> GetTable<T>(Database<T> database, string tableName) where T : Section, new();
    }
}
