#### New Blank Sheets Version 0.9-alpha

To import the macro into your Excel file, use (left) ALT + F11 or otherwise navigate to the VBAProject editor window for the workbook of interest. Then Use File -> Import File, then select NewBlankShts.bas. 

**You must have programmatic access to VBA Project enabled for this macro to work.**
To permit macro files this access, in your Excel workbook go to
File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> Trust access to the VBA project object model.

It is not recommended to use keyboard shortcuts to run this macro. Instead, assign the macro to a button in the Excel workbook. Look up tutorials online to help with this process, if needed.

The macro will look for very specific sheets when deciding what to copy. There are 3 requirements:

1. The worksheets' Names must start with "Blank " followed by whatever you like* 
  * *Note: if they are named exactly the same this will cause problems when re-naming the copied sheets, so be sure that the sheets' names all differ after 15 characters. The newly copied Blank sheets will be renamed with a timestamp to avoid naming conflicts with existing "working" sheets; this may truncate part of the name if your Blank sheets have long names.
2. The worksheets' VBAProject Excel Object module must have a CodeName starting with "Blank" followed immediately by sequentially numbering, with a leading "0" in front of one-digit numbers 
  * Blank01
  * Blank02
  * ...
  * Blank99
  * Blank100
  * *Etc.*
3. The worksheets must be Hidden or *xlveryhidden*

Some notes when creating these Blank sheets:
* When you are creating these sheets, take care that the formulas reference the sheets you want to reference.
* If formulas on your "working" (non-Blank) sheets change, be sure you change the corresponding formulas on your blank sheets.

At the end of the macro, all copied Blank sheets will have their VBAProject Module names renamed with the Excel-standard "Sheet" and the default sequential numbering scheme described above. The copied sheets will, of course, also be made visible.

This macro uses `ApplicationScreenUpdating = False` and `ApplicationScreenUpdating = True` to speed up the code, thus the user will not be able to see changes until after the conclusion of the macro.
