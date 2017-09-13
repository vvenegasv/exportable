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
    public class TestExcelXLSXImport
    {
        private const string DUMMY_PERSON = "00_DummyPerson.xlsx";
        private const string DUMMY_PERSON_DEFAULTS = "00_DummyPersonWithEmptyValues.xlsx";
        private const string MULTIPLE_DATA = "00_MultipleData.xlsx";
        private const string EMPTY_ROWS = "00_EmptyRows.xlsx";

        public TestExcelXLSXImport()
        {
            FileHelper.CreateResources();
        }

        [TestMethod]
        public void TestWithAttribute()
        {
            IImportEngine engine = new ExcelImportEngine();
            var key = engine.AddContainer<DummyPersonWithAttributes>();
            engine.SetDocument(PathConfig.BASE_PATH + DUMMY_PERSON);
            var data = engine.GetList<DummyPersonWithAttributes>(key);

            if (!data.Any() || data.Count != 30)
                throw new Exception("No se pudieron leer los registros del documento DummySimplePerson.xlsx");
        }

        [TestMethod]
        public void TestWithoutAttribute()
        {
            IImportEngine engine = new ExcelImportEngine();
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
            var key = engine.AsExcel().AddContainer<DummyPerson>("Dummy People", 1);
            engine.SetDocument(PathConfig.BASE_PATH + DUMMY_PERSON_DEFAULTS);
            var data = engine.GetList<DummyPersonWithAttributesAndDefaultValues>(key);

            if (!data.Any() || data.Count != 30)
                throw new Exception("No se pudieron leer los registros del documento " + DUMMY_PERSON_DEFAULTS);
        }

        [TestMethod]
        public void TestWithParallelReader()
        {
            IImportEngine engine = new ExcelImportEngine();
            engine.AsExcel().SetFormat(ExcelVersion.XLSX);
            engine.SetDocument(PathConfig.BASE_PATH + MULTIPLE_DATA);
            var resetEvents = new List<ManualResetEvent>();
            var countItemsInSheets = new List<int>();

            foreach (var index in new string[] { "1", "2" })
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
            var key = engine.AsExcel().AddContainer<DummyPerson>("Hoja1", 1);
            engine.SetDocument(PathConfig.BASE_PATH + EMPTY_ROWS);
            var data = engine.GetList<DummyPersonWithAttributesAndDefaultValues>(key);

            if (!data.Any() || data.Count != 33)
                throw new Exception("No se pudieron leer los registros del documento " + EMPTY_ROWS);
        }
    }
}

