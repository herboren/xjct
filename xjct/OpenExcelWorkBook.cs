using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;

namespace xjct
{
    public class OpenExcelWorkBook : IDisposable
    {        
        public Application xlApplication;
        private Workbook _workbook { get; set; }
        public string _filePath { get; private set; } = string.Empty;


        /// <summary>
        /// Get application configuration from ConfigureExcelApplication class.
        /// </summary>
        /// <param name="application"></param>
        /// <param name="filename"></param>
        public OpenExcelWorkBook(Application application, string filename)
        {
            this.xlApplication = application;
            this._filePath = filename;
        }

        public bool OpenExcelWorkBooks()
        {
            try
            {
                _workbook = xlApplication.Workbooks.Open(_filePath);

            }
            catch
            {
                throw new WhiteListException("An error occured while trying to open the workbook.\n" +
                    "Ensure the datatable is not damaged or present.");
            }
            finally
            {
                if (_workbook != null) { Marshal.ReleaseComObject(_workbook); }
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {

            }
        }
    }
}
