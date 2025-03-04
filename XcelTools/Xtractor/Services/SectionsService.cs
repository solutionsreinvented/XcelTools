using XcelTools.Xtractor.Enums;
using XcelTools.Xtractor.Interfaces;

using XcelTools.Xtractor.Models.Sections;

using ReInvented.DataAccess;
using ReInvented.DataAccess.Interfaces;

using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ReInvented.Sections.Domain.Repositories;
using ReInvented.Sections.Domain.Models;

using Excel = Microsoft.Office.Interop.Excel;
using System.Windows;
using XcelTools.Services;

namespace XcelTools.Xtractor.Services
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class SectionsService : ISectionsService
    {
        #region Private Fields

        private static SectionsService _instance;
        private readonly SectionsRepository _repository;
        private readonly SectionsLibrary _library;
        private static readonly SecurityService _security = new SecurityService();

        #endregion

        #region Private Constructor

        private SectionsService()
        {
            _repository = SectionsRepository.Instance;
            _library = _repository.GetSectionsLibrary();
        }

        #endregion

        #region Instance Creator

        public static SectionsService Instance => _instance ?? (_instance = new SectionsService());

        #endregion

        [ComVisible(true)]
        public void PopulateSectionClassifications(string dbName, string shape, Excel.Range rngTarget)
        {
            Excel.Worksheet worksheet = rngTarget.Worksheet;
            string wsName = worksheet.Name;
            _security.UnlockWorksheet(worksheet);
            //xSec.DisableSecurity(wsName)

            try
            {
                Database database = _library.Databases.FirstOrDefault(db => db.Name == dbName);
                SectionShape selectedShape = database.SectionShapes.FirstOrDefault(s => s.Shape == shape);
                string listItems = string.Join(",", selectedShape.Classifications.Select(c => c.Category).ToArray());
                ExcelCommonServices.PopulateListInACell(rngTarget, listItems);

                int sRow = rngTarget.Row;
                int sCol = rngTarget.Column;
                worksheet.Cells[sRow, sCol].Value = selectedShape.Classifications.FirstOrDefault().Category;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error encountered! Refer details below for further information. {Environment.NewLine}{ex.Message}", "Populate Section Types");
            }
            finally
            {
                _security.LockWorksheet(worksheet);
                //xSec.EnableSecurity(wsName)
            }
        }

        //[ComVisible(true)]
        //public void PopulateSectionProfiles(string dbName, string shape, string classification, Excel.Range rngTarget, Excel.Range rngPopulation)
        //{
        //    Excel.Worksheet worksheet = rngTarget.Worksheet;
        //    string wsName = worksheet.Name;
        //    _security.UnlockWorksheet(worksheet);
        //    //xSec.DisableSecurity(wsName)

        //    try
        //    {
        //        Database database = _library.Databases.FirstOrDefault(db => db.Name == dbName);
        //        SectionShape selectedShape = database.SectionShapes.FirstOrDefault(s => s.Shape == shape);
        //        Classification selectedClassification = selectedShape.Classifications.FirstOrDefault(c => c.Category == classification);

        //        string listItems = string.Join(",", selectedClassification.Sections.Select(s => s.Designation).ToArray());
        //        ExcelCommonServices.PopulateListInACell(rngTarget, listItems);

        //        int sRow = rngTarget.Row;
        //        int sCol = rngTarget.Column;
        //        worksheet.Cells[sRow, sCol].Value = selectedClassification.Sections.FirstOrDefault().Designation;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"An error encountered! Refer details below for further information. {Environment.NewLine}{ex.Message}", "Populate Section Types");
        //    }
        //    finally
        //    {
        //        _security.LockWorksheet(worksheet);
        //        //xSec.EnableSecurity(wsName)
        //    }
        //}

        [ComVisible(true)]
        public void PopulateSectionProfiles(string dbName, string shape, string classification, Excel.Range rngTarget, Excel.Range rngPopulation)
        {
            Excel.Worksheet wsTarget = rngTarget.Worksheet;
            Excel.Worksheet wsPopulation = rngPopulation.Worksheet;

            Excel.Application xlApp = wsTarget.Application;
            Excel.Workbook workbook = rngTarget.Parent as Excel.Workbook;

            string wsName = wsTarget.Name;
            _security.UnlockWorksheet(wsTarget);
            //xSec.DisableSecurity(wsName)

            try
            {
                Database database = _library.Databases.FirstOrDefault(db => db.Name == dbName);
                SectionShape selectedShape = database.SectionShapes.FirstOrDefault(s => s.Shape == shape);
                Classification selectedClassification = selectedShape.Classifications.FirstOrDefault(c => c.Category == classification);

                string[] profiles = selectedClassification.Sections.Select(s => s.Designation).ToArray();

                rngPopulation.Font.Name = "Tahoma";
                rngPopulation.Font.Size = 8;

                var sRow = rngPopulation.Row;
                var sCol = rngPopulation.Column;

                for (int i = 0; i < profiles.Length; i++)
                {
                    wsPopulation.Cells[sRow + i, sCol].Value = profiles[i];
                }

                var listItems = "=" + wsPopulation.Name + "!" + wsPopulation.Range[wsPopulation.Cells[sRow, sCol], wsPopulation.Cells[profiles.Length, sCol]].Address;

                ExcelCommonServices.PopulateListInACell(rngTarget, listItems);
                rngTarget.Value2 = profiles.FirstOrDefault();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error encountered! Refer details below for further information. {Environment.NewLine}{ex.Message}", "Populate Section Types");
            }
            finally
            {
                _security.LockWorksheet(wsTarget);
                //xSec.EnableSecurity(wsName)
            }
        }

        #region Public Functions

        [ComVisible(true)]
        public IClassificationInfo GetClassification(ShapeInfo shapeInfo, string category)
        {
            return shapeInfo.Classifications.FirstOrDefault(c => c.Category == category);
        }

        [ComVisible(true)]
        public IShapeInfo GetShapeInfo(DatabaseInfo databaseInfo, string shape)
        {
            return databaseInfo.SectionShapes.FirstOrDefault(s => s.Shape == (Shape)Enum.Parse(typeof(Shape), shape));
        }

        [ComVisible(true)]
        public IDatabaseInfo GetDatabaseInfo(string country)
        {
            return DatabaseInfo.Transform(_library.Databases.FirstOrDefault(d => d.Country == country));
        }

        //[ComVisible(true)]
        //public Catalogue GetSectionsCatalogue()
        //{
        //    string fileFullPath = Path.Combine(ServiceProvider.JsonSectionsDirectory, @"SectionsCatalogue.json");

        //    IDataSerializer<Catalogue> serializer = new JsonDataSerializer<Catalogue>();

        //    return serializer.Deserialize(fileFullPath);
        //}


        #endregion
    }
}
