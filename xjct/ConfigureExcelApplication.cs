using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;

namespace xjct
{
    public class ConfigureExcelApplication
    {
        static private Application _xlApplication = null;

        static OpenExcelWorkBook openExcelWorkBook;

        /// <summary>
        /// This will ensure the application does not appear or display alerts.
        /// Ultimately, we just want to configure the application enough to open
        /// the file, grab the data, and close the file, then release the object.
        /// </summary>
        /// <returns></returns>
        /// <exception cref="WhiteListException"></exception>
        static public bool ConfigureApplication(string filename)
        {
            try
            {
                // Configure Excel application object
                _xlApplication = new Application
                {
                    DisplayAlerts = false,  // Hide Alerts
                    Visible = false,        // Hide Excel Application 
                    ScreenUpdating = false, // Do not update visual elements, increases performance (large datasets) despite Visible = false.
                    EnableEvents = false,   // Do not execute workbook events
                    Calculation =           // Prevent automatic calculations when code is running.
                    XlCalculation.xlCalculationManual
                };

                // Pass this off to the next class for opening the workbook
                openExcelWorkBook = new OpenExcelWorkBook(_xlApplication, filename);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Could not configure application object.\n{ex.Message}");
                return false;
            }
            finally
            {
                if (_xlApplication != null)
                {
                    // Release comobject reference 
                    Marshal.ReleaseComObject(_xlApplication);
                }
            }
        }        
    }
}
