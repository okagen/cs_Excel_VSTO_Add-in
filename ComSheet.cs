using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using System.Runtime.InteropServices;


namespace CS_Excel_VSTO_Add_in
{
    class ComSheet : IDisposable
    {
        private Excel.Worksheet _activeSheet = null;
        private bool _disposed = false;

        private string _sheetTypeString = "";
        private string _sheetType = "";
        private string _sheetKind = "";
        private string _sheetVariation = "";

        public bool IsSummarySheet { get; set; }
        public bool IsDetailSheet { get; set; }
        public string UserDomain { get; set; }

        public void InitSheet(object sh)
        {
            _activeSheet = (Excel.Worksheet)sh;

            // Determine if the active sheet is the sheet for the estimate.
            if (GetSheetTypeStringFromCell())
            {
                SetDomainValueOnActiveSheet();
                this.IsSummarySheet = _sheetKind == ConfigurationManager.AppSettings["key_SummarySheet"];
                this.IsDetailSheet = _sheetKind == ConfigurationManager.AppSettings["key_DetailSheet"];
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing && _activeSheet != null)
                {
                        Marshal.ReleaseComObject(_activeSheet);
                        _activeSheet = null;
                }

                _disposed = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~ComSheet()
        {
            Dispose(false);
        }

        /// <summary>
        /// Set user domain on the active sheet.
        /// </summary>
        /// <param name="sh">object of an active sheet</param>
        private void SetDomainValueOnActiveSheet()
        {
            UserDomain = Environment.GetEnvironmentVariable("USERDOMAIN");
            string addressKey = ConfigurationManager.AppSettings["address_Domain"] + _sheetType + _sheetKind;
            string addressVal = ConfigurationManager.AppSettings[addressKey];
            _activeSheet.Range[addressVal].Value = UserDomain;
        }

        /// <summary>
        /// Determine if the active sheet is the sheet for the estimate.
        /// </summary>
        /// <returns></returns>
        private bool GetSheetTypeStringFromCell()
        {
            string rowSheetTypeString = ConfigurationManager.AppSettings["row_SheetTypeString"];
            string headerSheetTypeString = ConfigurationManager.AppSettings["header_SheetTypeString"];

            if (!string.IsNullOrEmpty(rowSheetTypeString) && !string.IsNullOrEmpty(headerSheetTypeString))
            {
                Excel.Range searchRange = _activeSheet.Range[rowSheetTypeString];
                Excel.Range foundCell = searchRange.Find(headerSheetTypeString, Type.Missing,
                    Excel.XlFindLookIn.xlValues,Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows,
                    Excel.XlSearchDirection.xlNext, false, false, Type.Missing);

                if (foundCell != null)
                {
                    object sheetTypeString = foundCell.Offset[0, 1].Value;
                    
                    if (sheetTypeString != null && !string.IsNullOrEmpty(sheetTypeString.ToString()))
                    {
                        _sheetTypeString = sheetTypeString.ToString();

                        if (ParseSheetTypeString())
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Split the sheet type string using "-", and assign "Type", "Kind", and "Variation" as member variables.
        /// </summary>
        /// <returns></returns>
        private bool ParseSheetTypeString()
        {
            try
            {
                string sheetTypeStringDelimiter = ConfigurationManager.AppSettings["delimiter_SheetTypeString"];
                if (string.IsNullOrEmpty(sheetTypeStringDelimiter))
                {
                    return false;
                }

                string[] splitSheetTypeString = _sheetTypeString.Split(Convert.ToChar(sheetTypeStringDelimiter));
                const int ExpectedSplitCount = 3;
                if (splitSheetTypeString.Length < ExpectedSplitCount)
                {
                    return false;
                }

                _sheetType = splitSheetTypeString[0];
                _sheetKind = splitSheetTypeString[1];
                _sheetVariation = splitSheetTypeString[2];
                return true;
            }
            catch (IndexOutOfRangeException ex)
            {
                Console.WriteLine("An error occurred while parsing the sheet type string: " + ex.Message);
                return false;
            }
        }


    }
}
