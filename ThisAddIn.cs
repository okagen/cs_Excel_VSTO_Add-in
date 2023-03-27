using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

//CustomTaskPaneを利用するために追加
using Microsoft.Office.Tools;

//Configuration Managerを利用するために追加
using System.Configuration;

namespace CS_Excel_VSTO_Add_in
{
    public partial class ThisAddIn
    {
        // global変数
        public UserControl1 g_UserControl1 { get; private set; }
        public CustomTaskPane g_TaskPane { get; private set; }


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // カスタムタスクパネルをアドインに追加する
            g_UserControl1 = new UserControl1();
            g_TaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(g_UserControl1, "My Task Pane");

            // シートが切り替わったときのイベントハンドラーを設定
            ((Excel.AppEvents_Event)this.Application).SheetActivate += new Excel.AppEvents_SheetActivateEventHandler(OnSheetActivated);

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // イベントハンドラーを解除
            ((Excel.AppEvents_Event)this.Application).SheetActivate -= new Excel.AppEvents_SheetActivateEventHandler(OnSheetActivated);
        }

        private void OnSheetActivated(object sh)
        {

            ComSheet csh = new ComSheet();
            csh.InitSheet(sh);

            // リボンのGroup3の表示を切り替える
            Globals.Ribbons.Ribbon1.group3.Visible = csh.IsSummarySheet;

            // リボンのGroup4の表示を切り替える
            Globals.Ribbons.Ribbon1.group4.Visible = csh.IsDetailSheet;

        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
