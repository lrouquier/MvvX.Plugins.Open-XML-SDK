using System;
using System.IO;

namespace MvvX.Plugins.OpenXMLSDK.Excel
{
    public interface IExcelManager : IDisposable
    {
        #region Create / Open / Save
        string WorksheetName { get; set; }
        /// <summary>
        /// Create a new instance of the ExcelPackage class based on a stream
        /// </summary>
        /// <param name="newStream">The stream object can be empty or contain a package. The stream must be Read/Write</param>
        bool CreateDoc(Stream newStream);

        /// <summary>
        /// Create a new excel document based on a template
        /// </summary>
        /// <param name="newStream">The output stream. Must be an empty read/write stream.</param>
        /// <param name="templateStream">This stream is copied to the output stream at load</param>
        bool CreateDocFromTemplate(Stream newStream, Stream templateStream);

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
        bool CreateWorksheet(string name);

        /// <summary>
        /// Add value to a cell
        /// </summary>
		/// <param name="row">The row number in the worksheet</param>
		/// <param name="col">The column number in the worksheet</param>
        /// <param name="value">value of the cell</param>
        bool AddCell(int row, int col, string value);

        /// <summary>
        /// Add value to a cell
        /// </summary>
        /// <param name="cell">cell ex : A2</param>
        /// <param name="value">value of the cell</param>
        bool AddCell(string cell, string value);

        /// <summary>
        /// Add value to a cell
        /// </summary>
        /// <param name="cell">cell ex : A2</param>
        /// <param name="value">value of the cell</param>
        bool AddCell(string cell, int value);

        #endregion
    }
}
