using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using Infodinamica.Framework.Exportable.Attribute;

namespace Infodinamica.Framework.Exportable.Tools
{
    internal static class MetadataHelper
    {
        public static IList<string> GetHeadersName(Type type)
        {
            var headeNames = new List<string>();            
            foreach (var member in type.GetMembers())
            {
                if (member.MemberType == MemberTypes.Property)
                {
                    var hasProperty = false;
                    foreach (var att in member.GetCustomAttributes(true))
                    {
                        if (att.GetType() == typeof(ExportableAttribute))
                        {
                            hasProperty = true;
                            headeNames.Add((att as ExportableAttribute).GetHeaderName());
                        }
                    }
                    
                    //Si no tiene el atributo ExcelBuilderAttribute, se ingresará al final de la lista y el encabezado será el nombre de la propiedad
                    if (!hasProperty)
                    {
                        headeNames.Add(((PropertyInfo)member).Name);
                    }
                }
            }

            return headeNames;

        }

        public static IList<ExportableMetadata> GetExportableMetadas(Type type)
        {
            var exportableMetadatas = new List<ExportableMetadata>();
            var exportableWithoutMetadata = new List<ExportableMetadata>();
            
            foreach (var member in type.GetMembers())
            {
                if (member.MemberType == MemberTypes.Property)
                {
                    var hasExportableAttribute = false;
                    foreach (var att in member.GetCustomAttributes(true))
                    {
                        if (att.GetType() == typeof(ExportableAttribute))
                        {
                            hasExportableAttribute = true;
                            var exportableAttribute = (att as ExportableAttribute);
                            exportableMetadatas.Add(new ExportableMetadata(member.Name, exportableAttribute.GetPosition(), exportableAttribute.GetFormat(), exportableAttribute.GetTypeValue()));
                        }
                    }

                    //If it havent ExportableAttribute, it will be added to another list because they will be in the last records
                    if (!hasExportableAttribute)
                    {
                        exportableWithoutMetadata.Add(new ExportableMetadata(member.Name, 0, null, FieldValueType.Any));
                    }
                }
            }

            //get biggest position
            int index = 0;
            if (exportableMetadatas.Any())
                index = exportableMetadatas
                    .Select(exp => exp.Position)
                    .Max();

            //Add elements without ExportableAttribute to returning list
            exportableWithoutMetadata.ForEach(exp =>
            {
                index++;
                exp.Position = index;
                exportableMetadatas.Add(exp);
            });

            return exportableMetadatas;
        }

        public static Type GetGenericType(object value)
        {
            return value.GetType().GetGenericArguments()[0];
        }

        public static Array GetArrayData(object value)
        {
            var toArrayMethod = value.GetType().GetMethod("ToArray");
            Array stronglyTypedArray = (Array)toArrayMethod.Invoke(value, null);
            return stronglyTypedArray;
        }
    }
}
