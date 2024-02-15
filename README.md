
# Overview
The repository macros command is a powerful feature designed to enhance workflow efficiency and automate repetitive tasks within a software repository environment. By defining macros, users can create custom commands or sequences of actions to streamline operations, increase productivity, and maintain consistency across projects.

# Getting Started
Press Alt + F11 to open the Visual Basic for Applications (VBA) editor. Alternatively, you can go to the "Developer" tab on the Excel ribbon (if not visible, you may need to enable it in Excel options), and click on "Visual Basic" to open the VBA editor.

## Insert a New Module
In the VBA editor window, go to the "Insert" menu and choose "Module." This action will insert a new module where you can write or paste your VBA code.

# Command
### Command for DeleteAllWorkbookCharts
```
Sub DeleteAllWorkbookCharts()

Dim wk As Worksheet

For Each wk In Worksheets

    If wk.ChartObjects.Count > 0 Then
        wk.ChartObjects.Delete
    End If
    
Next wk

End Sub
```
### Command for Save Chart
```
Option Explicit
Sub SaveChartKX()
Dim Chart_Obj As ChartObject
Dim Image_Chart As Chart
Dim nmyFileName As String

Set Chart_Obj = Sheets("Sheet1").ChartObjects(1)
Set Image_Chart = Chart_Obj.Chart
nmyFileName = "Penjualan_Perkategori.png"

On Error Resume Next
Kill ThisWorkbook.Path & "\" & nmyFileName
On Error GoTo 0
Image_Chart.Export Filename:=ThisWorkbook.Path & "\" & nmyFileName, Filtername:="PNG"

End Sub


```
### Notes
- Sub SaveChartKX():

- Sub keyword indicates the start of a subroutine definition named SaveChartKX.
- SaveChartKX is the name of the subroutine, which you can change to better describe its functionality or purpose.
- 
- Set Chart_Obj = Sheets("Sheet1").ChartObjects(1):
- Set keyword assigns the value to the variable Chart_Obj.
- Sheets("Sheet1") refers to the worksheet named "Sheet1" where the chart is located.
- ChartObjects(1) refers to the first chart object in the specified worksheet.
- You can change "Sheet1" to the name of the worksheet where your chart is located.
- You can change the index number (1 in this case) to access a different chart object if there are multiple charts in the worksheet.

- nmyFileName = "Penjualan_Perkategori.png":

- Assigns the filename to the variable nmyFileName.
- "Penjualan_Perkategori.png" is the default filename for the chart image.
- You can change "Penjualan_Perkategori.png" to any valid filename with the desired file extension.

- Image_Chart.Export Filename:=ThisWorkbook.Path & "" & nmyFileName, Filtername:="PNG":

- Image_Chart.Export exports the chart as an image file.
- Filename:=ThisWorkbook.Path & "\" & nmyFileName specifies the full path and filename for the exported image file.
- Filtername:="PNG" specifies the file format for the exported image (in this case, PNG format).
- You can change the file format (e.g., JPG, GIF) or modify the filename/path if needed.

# Run the Macros
To run the macro, you have several options:
- You can run the macro directly from the VBA editor by positioning the cursor anywhere within the DeleteAllWorkbookCharts subroutine and clicking the "Run" button (green triangle) on the toolbar or by pressing F5.
- Alternatively, you can run the macro from Excel itself by going to the "Developer" tab (if not visible, enable it in Excel options), clicking on "Macros," selecting the macro name (DeleteAllWorkbookCharts), and clicking "Run."
