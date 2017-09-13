using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Exportable.Helpers
{
    /// <summary>
    /// Métodos para obtener información en tiempo de ejecución de una clase determinada
    /// </summary>
    public static class TypeMethods
    {
        /// <summary>
        /// Revisa si es numérico
        /// </summary>
        /// <param name="prop"></param>
        /// <returns>Retorna verdadero si es un número</returns>
        public static bool IsNumeric(this PropertyInfo prop)
        {
            if (
                prop.PropertyType == typeof(sbyte) ||
                prop.PropertyType == typeof(byte) ||
                prop.PropertyType == typeof(short) ||
                prop.PropertyType == typeof(ushort) ||
                prop.PropertyType == typeof(int) ||
                prop.PropertyType == typeof(uint) ||
                prop.PropertyType == typeof(long) ||
                prop.PropertyType == typeof(ulong) ||
                prop.PropertyType == typeof(double) ||
                prop.PropertyType == typeof(decimal) ||
                prop.PropertyType == typeof(SByte) ||
                prop.PropertyType == typeof(Byte) ||
                prop.PropertyType == typeof(Int16) ||
                prop.PropertyType == typeof(Int32) ||
                prop.PropertyType == typeof(Int64) ||
                prop.PropertyType == typeof(UInt16) ||
                prop.PropertyType == typeof(UInt32) ||
                prop.PropertyType == typeof(UInt64) ||
                prop.PropertyType == typeof(Double) ||
                prop.PropertyType == typeof(Decimal)
                )
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Revisa si es fecha
        /// </summary>
        /// <param name="prop">Propiedad a revisar</param>
        /// <returns>Retorna verdadero si es una fecha</returns>
        public static bool IsDate(this PropertyInfo prop)
        {
            return prop.PropertyType == typeof(DateTime);
        }

        /// <summary>
        /// Revisa si es TimeSpan
        /// </summary>
        /// <param name="prop">Propiedad a revisar</param>
        /// <returns>Retorna verdadero si es TimeSpan</returns>
        public static bool IsTime(this PropertyInfo prop)
        {
            return prop.PropertyType == typeof(TimeSpan);
        }

        /// <summary>
        /// Revisa si es date o TimeSpan
        /// </summary>
        /// <param name="prop">Propiedad a revisar</param>
        /// <returns>Retorna verdadero si es fecha o TimeSpan</returns>
        public static bool IsDateOrTime(this PropertyInfo prop)
        {
            return prop.IsDate() || prop.IsTime();
        }

        /// <summary>
        /// Revisa si es boolean
        /// </summary>
        /// <param name="prop">Propiedad a revisar</param>
        /// <returns>Retorna verdadero si es boolean</returns>
        public static bool IsBoolean(this PropertyInfo prop)
        {
            return prop.PropertyType == typeof(Boolean) || prop.PropertyType == typeof(bool);
        }

        /// <summary>
        /// Revisa si el campo acepta nulos
        /// </summary>
        /// <param name="prop">Propiedad a revisar</param>
        /// <returns>Retorna verdadero si acepta nulos</returns>
        public static bool IsNullable(this PropertyInfo prop)
        {
            return !prop.PropertyType.IsValueType || prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>);
        }
    }
}
