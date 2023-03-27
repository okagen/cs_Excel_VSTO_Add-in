using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;

namespace CS_Excel_VSTO_Add_in
{
    class ComSheet
    {
        public bool IsSummarySheet { get; set; }
        public bool IsDetailSheet { get; set; }
        public string UserDomain { get; set; }

        private String _shTypeString = "";
        private String _shSyle = "";
        private String _shKind = "";
        private String _shVariation = "";


        public void InitSheet(object sh)
        {
            Excel.Worksheet activeSheet = (Excel.Worksheet)sh;

            //Determine if the active sheet is the sheet for the estimate.
            if (HasSheetTypeString(activeSheet))
            {
                SetUserDomainOnActiveSheet(activeSheet);
                this.IsSummarySheet = _shKind == ConfigurationManager.AppSettings["shSummary"];
                this.IsDetailSheet = _shKind == ConfigurationManager.AppSettings["shDetail"];
            }
        }

        /// <summary>
        /// Set user domain on the active sheet.
        /// </summary>
        /// <param name="sh"></param>
        private void SetUserDomainOnActiveSheet(in Excel.Worksheet sh)
        {
            string colKey = ConfigurationManager.AppSettings["col_Domain"] + _shSyle + _shKind;
            string rowKey = ConfigurationManager.AppSettings["row_Domain"] + _shSyle + _shKind;
            int col = Convert.ToInt16(ConfigurationManager.AppSettings[colKey]);
            int row = Convert.ToInt16(ConfigurationManager.AppSettings[rowKey]);
            this.UserDomain = Environment.GetEnvironmentVariable("USERDOMAIN");
            sh.Cells[row, col] = this.UserDomain;
        }

        /// <summary>
        /// Determine if the active sheet is the sheet for the estimate.
        /// </summary>
        /// <param name="sh">object of an active sheet</param>
        /// <returns>true : the sheet is for the estimate.</returns>
        private bool HasSheetTypeString(in Excel.Worksheet sh)
        {
            Excel.Range searchRange = sh.Range[ConfigurationManager.AppSettings["row_shType"]];
            Excel.Range foundCell = searchRange.Find(ConfigurationManager.AppSettings["hdr_shTypeString"],
                Type.Missing, Excel.XlFindLookIn.xlValues,
                Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext,
                false, false, Type.Missing);

            if (foundCell != null) 
            {
                string shTypeString = foundCell.Offset[0, 1].Value;
                if (SplitSheetTypeString(shTypeString))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Separate the string representing the sheet type with "-", then  divide it into "Style", "Kind", and "Variation" and set it as a member variable.
        /// </summary>
        /// <param name="shTypeString">String as a sheet type.</param>
        /// <returns></returns>
        private bool SplitSheetTypeString(in string shTypeString)
        {
            try
            {
                _shTypeString = shTypeString;
                string[] tmp = _shTypeString.Split(Convert.ToChar(ConfigurationManager.AppSettings["hdr_shTypeString_Split"]));
                _shSyle = tmp[0];
                _shKind = tmp[1];
                _shVariation = tmp[2];
                return true;
            }
            catch (IndexOutOfRangeException ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }


    }


}
