using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Resources;
using Exportable.Attribute;
using Exportable.InternalModels;
using Exportable.Models;

namespace Exportable.Helpers
{
    internal static class MetadataHelper
    {
        public static IList<Metadata> GetImportableMetadatas(Type type)
        {
            var exportableMetadatas = new List<Metadata>();
            var exportableWithoutMetadata = new List<Metadata>();

            foreach (var member in type.GetMembers())
            {
                if (member.MemberType == MemberTypes.Property)
                {
                    var hasExportableAttribute = false;
                    foreach (var att in member.GetCustomAttributes(true))
                    {
                        if (att.GetType() == typeof(ImportableAttribute))
                        {
                            hasExportableAttribute = true;
                            var importableAttribute = (att as ImportableAttribute);
                            //exportableMetadatas.Add(new Metadata(member.Name, importableAttribute.GetPosition(), null, FieldValueType.Any));
                            exportableMetadatas.Add(new Metadata(member.Name, importableAttribute.Position, null, FieldValueType.Any, importableAttribute.DefaultForNullOrInvalidValues));
                        }
                    }

                    //If it havent ExportableAttribute, it will be added to another list because they will be in the last records
                    if (!hasExportableAttribute)
                    {
                        //exportableWithoutMetadata.Add(new Metadata(member.Name, 0, null, FieldValueType.Any));
                        exportableWithoutMetadata.Add(new Metadata(member.Name, 0, null, FieldValueType.Any, string.Empty));
                    }
                }
            }

            //get biggest position
            int index = 0;
            if (exportableMetadatas.Any())
            {
                index = exportableMetadatas.Select(exp => exp.Position).Max();
                index++;
            }

            //Add elements without ExportableAttribute to returning list
            exportableWithoutMetadata.ForEach(exp =>
            {
                exp.Position = index;
                exportableMetadatas.Add(exp);
                index++;
            });

            return exportableMetadatas;
        }

        public static RowStyle GetHeadersFormat(Type type)
        {
            RowStyle rowStyle;
            var defaultFontName = "Calibry";
            var defaultFontColor = "#FFFFFF";
            short defaultFontSize = 11;
            var defaultBorderColor = "#000000";
            var defaultBackColor = "#888888";
            foreach (var ct in type.GetCustomAttributes(true))
            {
                if (ct != null && ct.GetType() == typeof(ExportableExcelHeaderAttribute))
                {
                    ExportableExcelHeaderAttribute exportableAttribute = (ct as ExportableExcelHeaderAttribute);
                    rowStyle = new RowStyle();

                    if (!string.IsNullOrWhiteSpace(exportableAttribute.FontName))
                        rowStyle.FontName = exportableAttribute.FontName;
                    else
                        rowStyle.FontName = defaultFontName;
                    if (!string.IsNullOrWhiteSpace(exportableAttribute.FontColor))
                        rowStyle.FontColor = exportableAttribute.FontColor;
                    else
                        rowStyle.FontColor = defaultFontColor;
                    if ((exportableAttribute.FontSize) > 0)
                        rowStyle.FontSize = exportableAttribute.FontSize;
                    else
                        rowStyle.FontSize = defaultFontSize;
                    if (!string.IsNullOrWhiteSpace(exportableAttribute.BorderColor))
                        rowStyle.BorderColor = exportableAttribute.BorderColor;
                    else
                        rowStyle.BorderColor = defaultBorderColor;
                    if (!string.IsNullOrWhiteSpace(exportableAttribute.BackColor))
                        rowStyle.BackColor = exportableAttribute.BackColor;
                    else
                        rowStyle.BackColor = defaultBackColor;
                    return rowStyle;
                }

            }

            return new RowStyle(defaultFontColor, defaultFontName, defaultFontSize, defaultBorderColor, defaultBackColor);
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

        public static string GetSheetNameFromAttribute(Type type)
        {
            foreach (var ct in type.GetCustomAttributes(true))
            {
                if (ct != null && ct.GetType() == typeof (ImportableExcelHeaderAttribute))
                {
                    return (ct as ImportableExcelHeaderAttribute).GetSheetName();
                }
                    
            }

            return null;
        }

        public static int GetFirstRowWithDataFromAttribute(Type type)
        {
            foreach (var ct in type.GetCustomAttributes(true))
            {
                if (ct != null && ct.GetType() == typeof(ImportableExcelHeaderAttribute))
                {
                    return (ct as ImportableExcelHeaderAttribute).GetFirstRowWithData();
                }

            }

            return -1;
        }

    }
}
