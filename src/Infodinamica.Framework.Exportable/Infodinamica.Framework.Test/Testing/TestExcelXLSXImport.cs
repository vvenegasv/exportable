﻿using System;
using System.Linq;
using Infodinamica.Framework.Exportable.Engines;
using Infodinamica.Framework.Exportable.Engines.Excel;
using Infodinamica.Framework.Test.Tools;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Infodinamica.Framework.Test.Testing
{
    [TestClass]
    public class TestExcelXLSXImport
    {
        private const string DUMMY_PERSON = "00_DummyPerson.xlsx";
        private const string DUMMY_PERSON_DEFAULTS = "00_DummyPersonWithEmptyValues.xlsx";

        [TestMethod]
        public void TestWithAttribute()
        {
            IImportEngine engine = new ExcelImportEngine();
            engine.AddContainer<DummyPersonWithAttributes>("1");
            engine.SetDocument(PathConfig.BASE_PATH + DUMMY_PERSON);
            var data = engine.GetList<DummyPersonWithAttributes>("1");

            if (!data.Any() || data.Count != 30)
                throw new Exception("No se pudieron leer los registros del documento DummySimplePerson.xlsx");
        }

        [TestMethod]
        public void TestWithoutAttribute()
        {
            IImportEngine engine = new ExcelImportEngine();
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
            engine.AsExcel().AddContainer<DummyPerson>("1", "Dummy People", 1);
            engine.SetDocument(PathConfig.BASE_PATH + DUMMY_PERSON_DEFAULTS);
            var data = engine.GetList<DummyPersonWithAttributesAndDefaultValues>("1");

            if (!data.Any() || data.Count != 30)
                throw new Exception("No se pudieron leer los registros del documento " + DUMMY_PERSON_DEFAULTS);
        }
    }
}

