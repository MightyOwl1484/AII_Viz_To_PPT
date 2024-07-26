# AII_Viz_To_PPT
Automates the transfer of slides from Excel to PowerPoint

VBA Script: Export Charts to PowerPoint
This VBA script automates the process of exporting charts from an Excel workbook to a PowerPoint presentation. It copies charts from each worksheet in the Excel workbook and pastes them into corresponding slides in a PowerPoint presentation, maintaining the layout and formatting.

Prerequisites
Before running the script, ensure you have the following:

Microsoft Excel with VBA enabled.
Microsoft PowerPoint installed.
The source Excel workbook named viz.xlsx.
A PowerPoint template named N980L - Aircraft Inventory - AII Viz.ppt.
Script Overview
The script performs the following steps:

Opens the PowerPoint application and the specified PowerPoint template.
Counts the number of worksheets in the source Excel workbook.
Duplicates a specific slide in the PowerPoint template for each worksheet.
Loops through each worksheet, copies the first chart found, and pastes it into the corresponding slide in the PowerPoint presentation.
Positions and sizes the pasted chart to fit the slide layout.
Saves the updated PowerPoint presentation.
Closes the PowerPoint application.
Usage
Ensure the source Excel workbook (viz.xlsx) and the PowerPoint template (N980L - Aircraft Inventory - AII Viz.ppt) are available in the specified paths.
Open the source Excel workbook in Excel.
Open the VBA editor (Alt + F11).
Insert a new module and paste the script into the module.
Run the script by pressing F5 or by using the "Run" button in the VBA editor.
Detailed Function Description
Sub ExportChartsToPowerPoint()
This is the main subroutine that handles the entire process.

Variables and Objects
pptApp: Object for the PowerPoint application.
pptPres: Object for the PowerPoint presentation.
pptSlide: Object for the PowerPoint slide.
ws: Worksheet object for the current worksheet being processed.
chartObj: ChartObject variable for the chart in the worksheet.
destWB: Workbook object for the destination Excel workbook.
pptTemplate: String variable for the path to the PowerPoint template.
pptSlideCount: Integer variable to store the number of slides/worksheets.
chartFound: Boolean variable to check if a chart is found in the worksheet.
pastedShape: Object for the pasted chart shape in PowerPoint.
i: Integer variable for loop iterations.
retryCount: Integer variable for retry attempts when copying the chart.
Constants
chartWidth: Width of the chart in points.
chartHeight: Height of the chart in points.
chartLeft: Left position of the chart in points.
chartTop: Top position of the chart in points.
Steps
Define Chart Size and Position:

vba
Copy code
Const chartWidth As Single = 351.36 ' 4.88 inches
Const chartHeight As Single = 200.16 ' 2.78 inches
Const chartLeft As Single = 360 ' 5 inches
Const chartTop As Single = 89.28 ' 1.24 inches
Path to the PowerPoint Template:

vba
Copy code
pptTemplate = "C:\Users\dave2\OneDrive\Desktop\N980L - Aircraft Inventory - AII Viz.ppt"
Open PowerPoint and the Template:

vba
Copy code
Set pptApp = CreateObject("PowerPoint.Application")
pptApp.Visible = True
Set pptPres = pptApp.Presentations.Open(pptTemplate)
Count the Number of Worksheets:

vba
Copy code
Set destWB = Workbooks("viz.xlsx")
pptSlideCount = destWB.Worksheets.Count
Duplicate Slides:

vba
Copy code
For i = 1 To pptSlideCount - 1
    pptPres.Slides(2).Duplicate
Next i
Loop Through Worksheets and Copy Charts:

vba
Copy code
For i = 1 To pptSlideCount
    Set ws = destWB.Worksheets(i)
    chartFound = False

    If ws.ChartObjects.Count > 0 Then
        Set chartObj = ws.ChartObjects(1)
        chartFound = True
    End If

    If chartFound Then
        retryCount = 0
        On Error Resume Next
        Do
            chartObj.Chart.CopyPicture
            DoEvents
            retryCount = retryCount + 1
        Loop Until retryCount = 3 Or Err.Number = 0
        On Error GoTo 0

        If Err.Number <> 0 Then
            MsgBox "Failed to copy chart from worksheet: " & ws.Name, vbExclamation
            Err.Clear
        Else
            Set pptSlide = pptPres.Slides(i + 1)
            On Error Resume Next
            pptSlide.Shapes.Paste
            On Error GoTo 0

            If pptSlide.Shapes.Count > 0 Then
                Set pastedShape = pptSlide.Shapes(pptSlide.Shapes.Count)
                With pastedShape
                    .LockAspectRatio = msoFalse
                    .Width = chartWidth
                    .Height = chartHeight
                    .Left = chartLeft
                    .Top = chartTop
                    .Rotation = 0
                End With
            Else
                MsgBox "Failed to paste chart from worksheet: " & ws.Name, vbExclamation
            End If
        End If
    End If
Next i
Save the Updated PowerPoint Presentation:

vba
Copy code
On Error Resume Next
pptPres.SaveAs "C:\Users\dave2\OneDrive\Desktop\N980L - Aircraft Inventory - AII Viz_Updated.ppt"
If Err.Number <> 0 Then
    MsgBox "Failed to save the PowerPoint presentation.", vbExclamation
    Err.Clear
End If
On Error GoTo 0
Close PowerPoint and Clean Up:

vba
Copy code
On Error Resume Next
pptPres.Close
If Err.Number <> 0 Then
    MsgBox "Failed to close the PowerPoint presentation.", vbExclamation
    Err.Clear
End If
On Error GoTo 0

pptApp.Quit

Set pptSlide = Nothing
Set pptPres = Nothing
Set pptApp = Nothing
Set chartObj = Nothing
Set ws = Nothing
Set destWB = Nothing
Notes
Ensure the source Excel workbook and PowerPoint template are available and paths are correct.
The script assumes the first chart in each worksheet is the one to be copied.
Modify the script to fit any specific requirements or additional customization.
