# cs_Excel_VSTO_Add-in

## log
1. Add Ribbon and UserControl.
2. Place 2 buttons in the Ribbon and display CustomTaskPane that contains the UserControl when push the button. One button displays CustomTaskPane from the left, the other from the right.
    - Initialize UserControl and CustomTaskPane in the ThisAddIn_Startup method in ThisAddIn.cs, and display them within the button_Click method in Ribbon.cs.
3. Switch the Ribbon based on a specific string set in a cell on the Excel sheet.
