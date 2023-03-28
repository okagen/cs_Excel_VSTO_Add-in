# cs_Excel_VSTO_Add-in

## log
1. Add Ribbon and UserControl.
2. Place 2 buttons in the Ribbon and display CustomTaskPane that contains the UserControl when push the button. One button displays CustomTaskPane from the left, the other from the right.
    - Initialize UserControl and CustomTaskPane in the ThisAddIn_Startup method in ThisAddIn.cs, and display them within the button_Click method in Ribbon.cs.
3. Switch the Ribbon based on a string set in a specific cell on the Excel active sheet.
    - To toggle the ribbon, I defined a cell address and a specific string in the App.conig file.
4. Create a common class for manipulating sheets.
5. Improved the readability of the ComSheet class.
6. Add combobox and button controls on UserControl, then add some lines when the button is clicked. 
   - Adds new rows to the active worksheet by copying rows from the range above and inserting them below the current cell.
   - Adds the number of rows the user has selected in the combo box.