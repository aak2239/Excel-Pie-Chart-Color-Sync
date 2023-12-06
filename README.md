# Excel-Pie-Chart-Color-Sync
VBA Script for Excel to Automatically Color Pie Chart Slices Based on Cell Fill Colors.

This repository contains a VBA (Visual Basic for Applications) script designed for Microsoft Excel. The script, MatchAllPieSliceColors, is crafted to synchronize the colors of pie chart slices with the fill colors of referenced cells. This utility is particularly useful in scenarios where the visual representation of data in pie charts needs to match specific color coding defined in Excel cells.
Features
* Automatically matches pie chart slice colors to cell fill colors.
* Works across multiple worksheets within an Excel workbook.
* Specifically designed for pie charts but can be adapted for other chart types.
* User-friendly: Automatically skips non-pie charts and handles errors gracefully.
Getting Started
Prerequisites
* Microsoft Excel (the script is written for Excel VBA).
Installation
* 		Download the VBA script file (MatchAllPieSliceColors.bas) from the repository.
* 		Open your Excel workbook.
* 		Press Alt + F11 to open the VBA editor.
* 		In the VBA editor, go to File > Import File... and import the downloaded .bas file.
* 		Close the VBA editor and return to your Excel workbook.
Usage
To run the script:
* 		Ensure your workbook has pie charts and corresponding cells with desired fill colors.
* 		Open the VBA editor (Alt + F11).
* 		Locate the MatchAllPieSliceColors subroutine in the imported module.
* 		Run the subroutine (F5) while in the editor or assign it to a button/form control in Excel for easier access.
Contributing
Your contributions to enhance this script, fix bugs, or improve functionality are highly welcomed and appreciated. Please feel free to fork the repository, make your changes, and submit a pull request with a clear description of your modifications.
License
This project is open-sourced under the MIT License. See the LICENSE file for more details.
Additional Notes
* This script is tailored for Excel pie charts. If you need to adapt it for other chart types, some modifications might be necessary.
* The script assumes the workbook contains a worksheet named "Space Use". If your workbook has a different structure, you may need to adjust the script accordingly.
