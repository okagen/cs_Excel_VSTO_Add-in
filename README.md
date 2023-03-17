# cs_Excel_VSTO_Add-in

## log
1. Ribbonを追加、UserControlを追加。
2. Ribbon内にボタンを配置し、UserControlを含んだCustomTaskPaneを表示させる。Buttonは2つ、それぞれCustomTaskPaneを左から表示、右から表示する。
    - ThisAddIn.cs内の **ThisAddIn_Startupメソッド内でUserControlとCustomTaskPaneを初期化** し、Ribbon.cs内のbutton_Clickメソッド内で、表示させる。
3. Excelシートの特定のセルに設定された文字列によって、Ribbonを切り替える。
