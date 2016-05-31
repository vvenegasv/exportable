using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Infodinamica.Framework.Exportable.Engines;
using Infodinamica.Framework.Exportable.Engines.Excel;
using Infodinamica.Framework.Exportable.Tools;
using Infodinamica.Framework.Test.Tools;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Infodinamica.Framework.Test.Testing
{
    [TestClass]
    public class TestExcelXLSExport
    {
        private const string BASE_PATH = @"E:\Github\exportable\src\Infodinamica.Framework.Exportable\TestResults\TestFiles\";

        public TestExcelXLSExport()
        {
            if (!Directory.Exists(BASE_PATH))
                Directory.CreateDirectory(BASE_PATH);
        }

        [TestMethod]
        public void ExportWithPlainClass()
        {
            IList<DummyPerson> dummyPeople = new List<DummyPerson>();
            for (int index = 0; index < 30; index++)
            {
                dummyPeople.Add(DummyFactory.CreateDummyPerson());
            }

            IExportEngine engine = new ExcelExportEngine();
            engine.AddData(dummyPeople);
            var fileName = Guid.NewGuid().ToString() + "-plain-class.xls";
            var filePath = BASE_PATH + fileName;
            engine.Export(filePath);
        }

        [TestMethod]
        public void ExportWithPlainClassAndSheetName()
        {
            IList<DummyPerson> dummyPeople = new List<DummyPerson>();
            for (int index = 0; index < 30; index++)
            {
                dummyPeople.Add(DummyFactory.CreateDummyPerson());
            }

            IExcelExportEngine engine = new ExcelExportEngine();
            engine.AsExcel().AddData(dummyPeople, "Dummy People");
            var fileName = Guid.NewGuid().ToString() + "-plain-class-sheets-names.xls";
            var filePath = BASE_PATH + fileName;
            engine.Export(filePath);
        }

        [TestMethod]
        public void ExportWithAttributes()
        {
            IList<DummyPersonWithAttributes> dummyPeople = new List<DummyPersonWithAttributes>();
            for (int index = 0; index < 30; index++)
            {
                dummyPeople.Add(DummyFactory.CreateDummyPersonWithAttributes());
            }

            IExportEngine engine = new ExcelExportEngine();
            engine.AddData(dummyPeople);
            var fileName = Guid.NewGuid().ToString() + "-with-attributes.xls";
            var filePath = BASE_PATH + fileName;
            engine.Export(filePath);
        }

        [TestMethod]
        public void ExportWithAttributesAndSomeSheetsNames()
        {
            IList<DummyPersonWithAttributes> dummyPeopleSheet1 = new List<DummyPersonWithAttributes>();
            IList<DummyPersonWithAttributes> dummyPeopleSheet2 = new List<DummyPersonWithAttributes>();
            IList<DummyPersonWithAttributes> dummyPeopleSheet3 = new List<DummyPersonWithAttributes>();
            
            for (int index = 0; index < 30; index++)
            {
                dummyPeopleSheet1.Add(DummyFactory.CreateDummyPersonWithAttributes());
            }

            for (int index = 0; index < 50; index++)
            {
                dummyPeopleSheet2.Add(DummyFactory.CreateDummyPersonWithAttributes());
            }

            for (int index = 0; index < 25; index++)
            {
                dummyPeopleSheet3.Add(DummyFactory.CreateDummyPersonWithAttributes());
            }

            IExcelExportEngine engine = new ExcelExportEngine();
            engine.AsExcel().AddData(dummyPeopleSheet1, "Sheet Number 1");
            engine.AsExcel().AddData(dummyPeopleSheet2, "Another Sheet");
            engine.AsExcel().AddData(dummyPeopleSheet3, "Custom Name");
            var fileName = Guid.NewGuid().ToString() + "-with-attributes-some-sheets.xls";
            var filePath = BASE_PATH + fileName;
            engine.Export(filePath);
        }

        [TestMethod]
        public void ExportWithMixAttributes()
        {
            IList<DummyPersonWIthSomeAttributes> dummyPeople = new List<DummyPersonWIthSomeAttributes>();
            for (int index = 0; index < 30; index++)
            {
                dummyPeople.Add(DummyFactory.CreateDummyPersonWIthSomeAttributes());
            }

            IExportEngine engine = new ExcelExportEngine();
            engine.AddData(dummyPeople);
            var fileName = Guid.NewGuid().ToString() + "-mix-class.xls";
            var filePath = BASE_PATH + fileName;
            engine.Export(filePath);
        }

        [TestMethod]
        public void ExportWithHeaderAttributes()
        {
            IList<DummyPersonWithHeaderAttribute> dummyPeople = new List<DummyPersonWithHeaderAttribute>();
            for (int index = 0; index < 30; index++)
            {
                dummyPeople.Add(DummyFactory.CreateDummyPersonWIthHeaderAttributes());
            }

            IExportEngine engine = new ExcelExportEngine();
            engine.AddData(dummyPeople);
            var fileName = Guid.NewGuid().ToString() + "-header-class.xls";
            var filePath = BASE_PATH + fileName;
            engine.Export(filePath);
        }

        [TestMethod]
        public void ExportWithSomeHeaderAttributes()
        {
            IList<DummyPersonWithSomeHeaderAttribute> dummyPeople = new List<DummyPersonWithSomeHeaderAttribute>();
            for (int index = 0; index < 30; index++)
            {
                dummyPeople.Add(DummyFactory.CreateDummyPersonWIthSomeHeaderAttributes());
            }

            IExportEngine engine = new ExcelExportEngine();
            engine.AddData(dummyPeople);
            var fileName = Guid.NewGuid().ToString() + "-some-header-class.xls";
            var filePath = BASE_PATH + fileName;
            engine.Export(filePath);
        }

        [TestMethod]
        public void ExportWithResourceHeaderAttribute()
        {
            IList<DummyPersonWithAttributesAndResource> dummyPeople = new List<DummyPersonWithAttributesAndResource>();
            for (int index = 0; index < 30; index++)
            {
                dummyPeople.Add(DummyFactory.CreateDummyPersonWithAttributesAndResource());
            }

            IExportEngine engine = new ExcelExportEngine();
            engine.AddData(dummyPeople);
            var fileName = Guid.NewGuid().ToString() + "-some-header-class.xls";
            var filePath = BASE_PATH + fileName;
            engine.Export(filePath);
        }

        [TestMethod]
        public void ExportWithResourceHeaderAttributeWithOtherCulture()
        {
            IList<DummyPersonWithAttributesAndResource> dummyPeople = new List<DummyPersonWithAttributesAndResource>();
            for (int index = 0; index < 30; index++)
            {
                dummyPeople.Add(DummyFactory.CreateDummyPersonWithAttributesAndResource());
            }

            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-US");

            IExportEngine engine = new ExcelExportEngine();
            engine.AddData(dummyPeople);
            var fileName = Guid.NewGuid().ToString() + "-some-header-class.xls";
            var filePath = BASE_PATH + fileName;
            engine.Export(filePath);
        }

        [TestMethod]
        public void ExportIntensiveUsage()
        {
            IList<DummyPerson> dummyPeople = new List<DummyPerson>();
            for (int index = 0; index < 65000; index++)
            {
                dummyPeople.Add(DummyFactory.CreateDummyPerson());
            }

            IExportEngine engine = new ExcelExportEngine();
            engine.AddData(dummyPeople);
            var fileName = Guid.NewGuid().ToString() + "-intensive.xls";
            var filePath = BASE_PATH + fileName;
            engine.Export(filePath);
        }
    }
}
