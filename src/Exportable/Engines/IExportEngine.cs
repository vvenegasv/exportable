using System;
using System.Collections.Generic;
using System.IO;
using System.Linq.Expressions;
using Exportable.Engines.Excel;
using NPOI.Util;

namespace Exportable.Engines
{
    /// <summary>
    /// Interface to export
    /// </summary>
    public interface IExportEngine
    {
        /// <summary>
        /// Add dta to export
        /// </summary>
        /// <typeparam name="TModel">Model type of data to export</typeparam>
        /// <param name="data">Array with data to export</param>
        /// <returns>Key of the dataset to export</returns>
        string AddData<TModel>(IEnumerable<TModel> data) where TModel : class;

        /// <summary>
        /// Export the document
        /// </summary>
        /// <returns>MemoryStream with exported document</returns>
        MemoryStream Export();

        /// <summary>
        /// Export the document
        /// </summary>
        /// <param name="path">Path in wich the file will be save</param>
        void Export(string path);

        /// <summary>
        /// Run validations
        /// </summary>
        /// <returns>Dictionary with errors</returns>
        IDictionary<string, string> RunBusinessRules();

        /// <summary>
        /// Gets an interface to works with specialized excel method
        /// </summary>
        /// <returns></returns>
        IExcelExportEngine AsExcel();

        /// <summary>
        /// Set the ignore columns
        /// </summary>
        /// <typeparam name="TModel">Model type that has the column to ignore</typeparam>
        /// <param name="datasetKey">Key of the dataset</param>
        /// <param name="propertyExpression">Expression to ignore a property</param>
        void AddIgnoreColumns<TModel>(string datasetKey, Expression<Func<TModel, object>> propertyExpression)
            where TModel : class;

        /// <summary>
        /// Set the ignore columns
        /// </summary>
        /// <typeparam name="TModel">Model type that has the column to ignore</typeparam>
        /// <param name="datasetKey">Key of the dataset</param>
        /// <param name="propertyExpression">Expression to overrides a property name</param>
        /// <param name="name">New property name</param>
        void AddColumnsNames<TModel>(string datasetKey, Expression<Func<TModel, object>> propertyExpression,
            string name)
            where TModel : class;


    }
}
