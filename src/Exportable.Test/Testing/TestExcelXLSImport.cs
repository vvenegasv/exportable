using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Exportable.Engines;
using Exportable.Engines.Excel;
using Exportable.Test.Tools;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Exportable.Test.Testing
{
    [TestClass]
    public class TestExcelXLSImport
    {
        private const string DUMMY_PERSON = "00_DummyPerson.xls";
        private const string DUMMY_PERSON_DEFAULT = "00_DummyPersonWithEmptyValues.xls";
        private const string MULTIPLE_DATA = "00_MultipleData.xls";
        private const string EMPTY_ROWS = "00_EmptyRows.xls";

        public TestExcelXLSImport()
        {
            FileHelper.CreateResources();
        }

        [TestMethod]
        public void TestWithAttribute()
        {
            try
            {
                IImportEngine engine = new ExcelImportEngine();
                engine.AsExcel().SetFormat(ExcelVersion.XLS);
                var key =engine.AddContainer<DummyPersonWithAttributes>();
                engine.SetDocument(PathConfig.BASE_PATH + DUMMY_PERSON);
                var data = engine.GetList<DummyPersonWithAttributes>(key);

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
            var key = engine.AsExcel().AddContainer<DummyPerson>("Dummy People", 1);
            engine.SetDocument(PathConfig.BASE_PATH + DUMMY_PERSON);
            var data = engine.GetList<DummyPerson>(key);

            if (!data.Any() || data.Count != 30)
                throw new Exception("No se pudieron leer los registros del documento DummySimplePerson.xlsx");
        }

        [TestMethod]
        public void TestWithDefaults()
        {
            IImportEngine engine = new ExcelImportEngine();
            engine.AsExcel().SetFormat(ExcelVersion.XLS);
            var key = engine.AsExcel().AddContainer<DummyPerson>("Dummy People", 1);
            engine.SetDocument(PathConfig.BASE_PATH + DUMMY_PERSON_DEFAULT);
            var data = engine.GetList<DummyPersonWithAttributesAndDefaultValues>(key);

            if (!data.Any() || data.Count != 30)
                throw new Exception("No se pudieron leer los registros del documento " + DUMMY_PERSON_DEFAULT);
        }

        [TestMethod]
        public void TestWithParallelReader()
        {
            IImportEngine engine = new ExcelImportEngine();
            engine.AsExcel().SetFormat(ExcelVersion.XLS);
            engine.SetDocument(PathConfig.BASE_PATH + MULTIPLE_DATA);
            var resetEvents = new List<ManualResetEvent>();
            var countItemsInSheets = new List<int>();

            foreach (var index in new string[] {"1", "2"})
            {
                var evt = new ManualResetEvent(false);
                resetEvents.Add(evt);
                ThreadPool.QueueUserWorkItem(i =>
                {
                    var key = engine.AsExcel().AddContainer<DummyPerson>("Hoja" + index, 1);
                    var data = engine.GetList<DummyPersonWithAttributesAndDefaultValues>(key);
                    countItemsInSheets.Add(data.Count);
                    evt.Set();
                }, index);
            }

            foreach (var evt in resetEvents)
                evt.WaitOne();

            if (!countItemsInSheets.Any() || countItemsInSheets.Sum() != 60)
                throw new Exception("No se pudieron leer los registros en paralelo del documento " + MULTIPLE_DATA);
        }

        [TestMethod]
        public void TestWithEmptyRows()
        {
            IImportEngine engine = new ExcelImportEngine();
            engine.AsExcel().SetFormat(ExcelVersion.XLS);
            var key = engine.AsExcel().AddContainer<DummyPerson>("Hoja1", 1);
            engine.SetDocument(PathConfig.BASE_PATH + EMPTY_ROWS);
            var data = engine.GetList<DummyPersonWithAttributesAndDefaultValues>(key);

            if (!data.Any() || data.Count != 33)
                throw new Exception("No se pudieron leer los registros del documento " + EMPTY_ROWS);
        }
    }
}

