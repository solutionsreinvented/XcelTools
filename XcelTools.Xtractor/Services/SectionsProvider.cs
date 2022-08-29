using XcelTools.Xtractor.Enums;
using XcelTools.Xtractor.Interfaces;

using XcelTools.Xtractor.Models.Sections;

using ReInvented.DataAccess;
using ReInvented.DataAccess.Interfaces;

using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Reflection;

namespace XcelTools.Xtractor.Services
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class SectionsProvider : ISectionsProvider
    {
        #region Private Fields

        private static SectionsProvider _instance;

        #endregion

        #region Private Constructor

        private SectionsProvider()
        {

        }

        #endregion

        #region Instance Creator

        public static SectionsProvider Instance()
        {
            return _instance ?? new SectionsProvider();
        }

        #endregion

        #region Public Functions

        [ComVisible(true)]
        public Classification GetClassification(ShapeInfo shapeInfo, string category)
        {
            return shapeInfo.Classifications.FirstOrDefault(c => c.Category == category);
        }

        [ComVisible(true)]
        public ShapeInfo GetShapeInfo(DatabaseInfo databaseInfo, string shape)
        {
            return databaseInfo.SectionShapes.FirstOrDefault(s => s.Shape == (Shape)Enum.Parse(typeof(Shape), shape));
        }

        [ComVisible(true)]
        public DatabaseInfo GetDatabaseInfo(Catalogue catalogue, string country)
        {
            return catalogue.Databases.FirstOrDefault(d => d.Country == country);
        }

        [ComVisible(true)]
        public Catalogue GetSectionsCatalogue()
        {
            string fileFullPath = Path.Combine(ServiceProvider.JsonSectionsDirectory, @"SectionsCatalogue.json");

            IDataSerializer<Catalogue> serializer = new JsonDataSerializer<Catalogue>();

            return serializer.Deserialiaze(fileFullPath);
        }

        [ComVisible(true)]
        public Database<T> GetDatabase<T>(string fileName) where T : Section, new()
        {
            string fileFullPath = Path.Combine(ServiceProvider.JsonSectionsDirectory, fileName);

            IDataSerializer<Database<T>> databaseSerializer = new JsonDataSerializer<Database<T>>();

            return databaseSerializer.Deserialiaze(fileFullPath);
        }

        [ComVisible(true)]
        public Table<T> GetTable<T>(Database<T> database, string tableName) where T : Section, new()
        {
            return new Table<T>();
        }

        #endregion

        #region Private Helpers


        #endregion

    }
}
