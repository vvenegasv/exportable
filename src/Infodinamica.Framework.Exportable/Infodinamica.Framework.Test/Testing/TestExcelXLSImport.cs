using System;
using System.Linq;
using Infodinamica.Framework.Exportable.Engines;
using Infodinamica.Framework.Exportable.Engines.Excel;
using Infodinamica.Framework.Test.Tools;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Infodinamica.Framework.Test.Testing
{
    [TestClass]
    public class TestExcelXLSImport
    {
        private const string DUMMY_PERSON = "00_DummyPerson.xls";
        private const string DUMMY_PERSON_DEFAULT = "00_DummyPersonWithEmptyValues.xls";

        [TestMethod]
        public void TestWithAttribute()
        {
            try
            {
                IImportEngine engine = new ExcelImportEngine();
                engine.AsExcel().SetFormat(ExcelVersion.XLS);
                engine.AddContainer<DummyPersonWithAttributes>("1");
                engine.SetDocument(PathConfig.BASE_PATH + DUMMY_PERSON);
                var data = engine.GetList<DummyPersonWithAttributes>("1");

                if (!data.Any() || data.Count != 30)
                    throw new Exception("No se pudieron leer los registros del documento DummySimplePerson.xlsx");
            }
            catch (Exception ex)
            {
                throw new Exception("No se pudo completar la prueba. Revise el error interno", ex);
            }
        }

        [TestMethod]
        public void TestWithoutAttribute()
        {
            IImportEngine engine = new ExcelImportEngine();
            engine.AsExcel().SetFormat(ExcelVersion.XLS);
            engine.AsExcel().AddContainer<DummyPerson>("1", "Dummy People", 1);
            engine.SetDocument(PathConfig.BASE_PATH + DUMMY_PERSON);
            var data = engine.GetList<DummyPerson>("1");

            if (!data.Any() || data.Count != 30)
                throw new Exception("No se pudieron leer los registros del documento DummySimplePerson.xlsx");
        }

        [TestMethod]
        public void TestWithDefaults()
        {
            IImportEngine engine = new ExcelImportEngine();
            engine.AsExcel().SetFormat(ExcelVersion.XLS);
            engine.AsExcel().AddContainer<DummyPerson>("1", "Dummy People", 1);
            engine.SetDocument(PathConfig.BASE_PATH + DUMMY_PERSON_DEFAULT);
            var data = engine.GetList<DummyPersonWithAttributesAndDefaultValues>("1");

            if (!data.Any() || data.Count != 30)
                throw new Exception("No se pudieron leer los registros del documento " + DUMMY_PERSON_DEFAULT);
        }
    }
}

