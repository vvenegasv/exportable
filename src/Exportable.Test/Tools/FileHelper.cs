using System.IO;

namespace Exportable.Test.Tools
{
    internal static class FileHelper
    {
        public static void CreateResources()
        {
            if (!Directory.Exists(PathConfig.BASE_PATH))
                Directory.CreateDirectory(PathConfig.BASE_PATH);

            if (!File.Exists(string.Format(@"{0}\00_DummyPerson.xls", PathConfig.BASE_PATH)))
                File.WriteAllBytes(string.Format(@"{0}\00_DummyPerson.xls", PathConfig.BASE_PATH), Files.DummyPerson_xls);

            if (!File.Exists(string.Format(@"{0}\00_DummyPerson.xlsx", PathConfig.BASE_PATH)))
                File.WriteAllBytes(string.Format(@"{0}\00_DummyPerson.xlsx", PathConfig.BASE_PATH), Files.DummyPerson_xlsx);

            if (!File.Exists(string.Format(@"{0}\00_DummyPersonWithEmptyValues.xls", PathConfig.BASE_PATH)))
                File.WriteAllBytes(string.Format(@"{0}\00_DummyPersonWithEmptyValues.xls", PathConfig.BASE_PATH), Files.DummyPersonWithEmptyValues_xls);

            if (!File.Exists(string.Format(@"{0}\00_DummyPersonWithEmptyValues.xlsx", PathConfig.BASE_PATH)))
                File.WriteAllBytes(string.Format(@"{0}\00_DummyPersonWithEmptyValues.xlsx", PathConfig.BASE_PATH), Files.DummyPersonWithEmptyValues_xlsx);

            if (!File.Exists(string.Format(@"{0}\00_MultipleData.xls", PathConfig.BASE_PATH)))
                File.WriteAllBytes(string.Format(@"{0}\00_MultipleData.xls", PathConfig.BASE_PATH), Files.MultipleData_xls);

            if (!File.Exists(string.Format(@"{0}\00_MultipleData.xlsx", PathConfig.BASE_PATH)))
                File.WriteAllBytes(string.Format(@"{0}\00_MultipleData.xlsx", PathConfig.BASE_PATH), Files.MultipleData_xlsx);

            if (!File.Exists(string.Format(@"{0}\00_EmptyRows.xlsx", PathConfig.BASE_PATH)))
                File.WriteAllBytes(string.Format(@"{0}\00_EmptyRows.xlsx", PathConfig.BASE_PATH), Files.EmptyRows_xlsx);

            if (!File.Exists(string.Format(@"{0}\00_EmptyRows.xls", PathConfig.BASE_PATH)))
                File.WriteAllBytes(string.Format(@"{0}\00_EmptyRows.xls", PathConfig.BASE_PATH), Files.EmptyRows_xls);
        }
    }
}
