using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using MvvX.Plugins.OpenXMLSDK.Excel;
using OfficeOpenXml;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Excel
{
    public class ExcelManager : IExcelManager
    {

        #region Fields

        private ExcelPackage package = null;
        private ExcelWorksheet worksheet = null;

        private string worksheetName;
        public string WorksheetName
        {
            get
            {
                return this.worksheetName;
            }
            set
            {
                this.worksheetName = value;
            }
        }
        #endregion

        #region Dispose

        public void Dispose()
        {
            if (package != null)
                package.Dispose();
        }

        #endregion

        #region Create / Open / Save

        /// <summary>
        /// Create a new excel document
        /// </summary>
        /// <param name="newStream">The stream object can be empty or contain a package. The stream must be Read/Write</param>
        public bool CreateDoc(Stream newStream)
        {
            try
            {
                package = new ExcelPackage(newStream);
                return true;
            }
            catch (Exception e)
            {
                package = null;
                return false;
            }

        }

        /// <summary>
        /// Create a new excel document based on a template
        /// </summary>
        /// <param name="newStream">The output stream. Must be an empty read/write stream.</param>
        /// <param name="templateStream">This stream is copied to the output stream at load</param>
        public bool CreateDocFromTemplate(Stream newStream, Stream templateStream)
        {
            try
            {
                package = new ExcelPackage(newStream, templateStream);
                return true;
            }
            catch (Exception e)
            {
                package = null;
                return false;
            }
        }


        /// <summary>
        /// Saves all the components back into the package.
        /// We close the package after the save is done.
        /// </summary>
        public void Save(Stream OutputStream)
        {
            if(package != null)
            package.Save();
        }

        /// <summary>
        /// Saves all the components back into the package.
        /// We close the package after the save is done.
        /// </summary>
        public void Save()
        {
            if (package != null)
                package.Save();
        }

        #endregion

        #region Workbook

        /// <summary>
        /// Add a new worksheet
        /// </summary>
        /// <param name="name">name of the worksheet</param>
        /// <returns></returns>
        public bool CreateWorksheet(string name)
        {
            if (name == null)
                throw new ArgumentNullException("name must be not null");

            try
            {
                worksheet = package.Workbook.Worksheets.Add(name);
                return true;
            }
            catch (Exception e)
            {
                worksheet = null;
                return false;
            }

        }

        /// <summary>
        /// Add value to a cell
        /// </summary>
		/// <param name="row">The row number in the worksheet</param>
		/// <param name="col">The column number in the worksheet</param>
        /// <param name="value">value of the cell</param>
        public bool AddCell(int row, int col, string value)
        {
            if (row == 0)
                throw new ArgumentNullException("row must be not null");
            if (col == 0)
                throw new ArgumentNullException("col must be not null");
            if (value == null)
                throw new ArgumentNullException("value must be not null");

            try
            {
                if(worksheet == null)
                    worksheet = package.Workbook.Worksheets[WorksheetName];

                worksheet.Cells[row, col].Value = value;
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        /// <summary>
        /// Add value to a cell
        /// </summary>
        /// <param name="cell">cell ex : A2</param>
        /// <param name="value">value of the cell</param>
        public bool AddCell(string cell, string value)
        {
            if (cell == null)
                throw new ArgumentNullException("cell must be not null");
            if (value == null)
                throw new ArgumentNullException("value must be not null");

            try
            {
                if (worksheet == null)
                    worksheet = package.Workbook.Worksheets[WorksheetName];

                worksheet.Cells[cell].Value = value;
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        /// <summary>
        /// Add value to a cell
        /// </summary>
        /// <param name="cell">cell ex : A2</param>
        /// <param name="value">value of the cell</param>
        public bool AddCell(string cell, int value)
        {
            if (cell == null)
                throw new ArgumentNullException("cell must be not null");
            if (value == 0)
                throw new ArgumentNullException("value must be not null");

            try
            {
                if (worksheet == null)
                    worksheet = package.Workbook.Worksheets[WorksheetName];

                worksheet.Cells[cell].Value = value;
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        #endregion
    }
}
