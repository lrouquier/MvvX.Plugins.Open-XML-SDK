using System;
using System.IO;

namespace MvvX.Plugins.OpenXMLSDK.Excel
{
    public interface IExcelManager : IDisposable
    {
        #region Create / Open / Save

        /// <summary>
        /// Create a new instance of the ExcelPackage class based on a stream
        /// </summary>
        /// <param name="newStream">The stream object can be empty or contain a package. The stream must be Read/Write</param>
        bool CreateDoc(Stream newStream);

        /// <summary>
        /// Saves all the components back into the package.
        /// We close the package after the save is done.
        /// </summary>
        void Save(Stream OutputStream);

        /// <summary>
        /// Saves all the components back into the package.
        /// We close the package after the save is done.
        /// </summary>
        void Save();

        #endregion

        #region Workbook

        /// <summary>
        /// Add a new worksheet
        /// </summary>
        /// <param name="name">name of the worksheet</param>
        /// <returns></returns>
        bool CreateWorksheet(string name);

        #endregion
    }
}
