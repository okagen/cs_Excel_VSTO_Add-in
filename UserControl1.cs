using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;


namespace CS_Excel_VSTO_Add_in
{
    public partial class UserControl1 : UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
            {
                // Get number of rows to add.
                int intValue = Convert.ToInt32(comboBox1.SelectedItem);

                // 明細行のコピー元を取得
                Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet;
                Excel.Range sourceRange = worksheet.Range[ConfigurationManager.AppSettings["range_DetailLineToCopy"]];

                // CurrentCellを取得すし、A列までオフセット
                Excel.Range currentCell = worksheet.Application.ActiveCell;
                Excel.Range offsetCell = currentCell.Offset[0, -currentCell.Column + 1];

                // A列までオフセットしたCurrentCellの下にComboboxで選択された行数追加する
                Excel.Range insertRange = offsetCell.Resize[intValue, sourceRange.Columns.Count];
                insertRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                sourceRange.Copy(insertRange);
            }
            else
            {
                // intValueが取得できなかった場合の処理を記述する
                MessageBox.Show("Select number of rows to add.");
            }
        }
    }
}
