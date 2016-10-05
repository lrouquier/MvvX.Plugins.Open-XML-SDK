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
        /// Create a new instance of the ExcelPackage class based on a stream
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
                package.Workbook.Worksheets.Add(name);
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
