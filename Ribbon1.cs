using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

//CustomTaskPaneを利用するために追加
using Microsoft.Office.Tools;

namespace CS_Excel_VSTO_Add_in
{
    public partial class Ribbon1
    {
        private const int TASK_PANE_WIDTH = 600;
        private const string TASK_PANE_TITLE = "My Task Pane";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            label1.Label = "Globals.ThisAddIn.Application.UserName [" + Globals.ThisAddIn.Application.UserName + " ]";
            label2.Label = "環境変数 USERNAME [ " + Environment.GetEnvironmentVariable("username") + " ]";
            label3.Label = "環境変数 USERDOMAIN [ " + Environment.GetEnvironmentVariable("USERDOMAIN") + " ]";
        }

        private void button_FromTheLeft_Click(object sender, RibbonControlEventArgs e)
        {
            UserControl1 g_UC = Globals.ThisAddIn.g_UserControl1;
            CustomTaskPane g_TP = Globals.ThisAddIn.g_TaskPane;

            if (g_TP != null)
            {
                if (g_TP.Visible)
                {
                    g_TP.Visible = false;
                }
                else
                {
                    g_TP.Width = TASK_PANE_WIDTH;
                    g_TP.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
                    g_TP.Visible = true;
                }
            }
            else
            {
                // Common Task Paneがまだ作成されていない場合は作成する
                g_TP = Globals.ThisAddIn.CustomTaskPanes.Add(g_UC, TASK_PANE_TITLE);
            }
        }

        private void button_FromTheRight_Click(object sender, RibbonControlEventArgs e)
        {
            UserControl1 g_UC = Globals.ThisAddIn.g_UserControl1;
            CustomTaskPane g_TP = Globals.ThisAddIn.g_TaskPane;

            if (g_TP != null)
            {
                if (g_TP.Visible)
                {
                    g_TP.Visible = false;
                }
                else
                {
                    g_TP.Width = TASK_PANE_WIDTH;
                    g_TP.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                    g_TP.Visible = true;
                }
            }
            else
            {
                // Common Task Paneがまだ作成されていない場合は作成する
                g_TP = Globals.ThisAddIn.CustomTaskPanes.Add(g_UC, TASK_PANE_TITLE);
            }

        }
    }
}
