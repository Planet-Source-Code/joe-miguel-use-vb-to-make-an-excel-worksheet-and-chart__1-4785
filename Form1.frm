VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' By Joe Miguel   joe_miguel@hotmail.com
'
' Remember - to speed this up, do NOT make excel visible until you are finished.
' Making it visible from the beginning, excel is redrawn after every instruction.
' I did not do it this way because i want to show the process.

Private Sub Command1_Click()
Dim varNum As Long

Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.Worksheet

'Start the excel COM and make it visible.
Set objExcel = GetObject("", "excel.application")
'Set objExcel = excel.Application ' Seems to cause a memory leak
    objExcel.Visible = True
    
'Start a workbook.
Set objWorkbook = objExcel.Workbooks.Add

'Turn off the alerts, otherwise user will have to confirm my actions.
    objExcel.DisplayAlerts = False

'Depending on the users excel's settings, there could be many worksheet when starting a workbook.
'Ensure there is only one worksheet.
Do While objWorkbook.Worksheets.Count > 1
    Set objWorksheet = objWorkbook.Worksheets.Item(objWorkbook.Worksheets.Count)
    objWorksheet.Delete
Loop

'Set objWorksheet to the remaining worksheet.
Set objWorksheet = ActiveSheet

'Rename the sheet to Results.
    objWorksheet.Name = "Results"

'Headers
    objWorksheet.Cells(1, 1) = "Blah Blah Blah Analytic Labs"
    objWorksheet.Cells(1, 1).Font.Bold = True
    objWorksheet.Cells(2, 1) = "Experiment Name"
    objWorksheet.Cells(2, 1).Font.Bold = True
    objWorksheet.Cells(2, 3) = "Trial Number"
    objWorksheet.Cells(2, 3).Font.Bold = True
    objWorksheet.Cells(2, 5) = "Batch Number"
    objWorksheet.Cells(2, 5).Font.Bold = True
    objWorksheet.Cells(3, 1) = " " & Now
    objWorksheet.Cells(3, 1).Font.Bold = True
    
'Results
    objWorksheet.Cells(5, 1) = "Results"
    objWorksheet.Cells(5, 1).Font.Bold = True
    
'General info
    objWorksheet.Cells(8, 2) = "Number of Samples:"
    objWorksheet.Cells(9, 2) = "Sample Amount (ug):"
    objWorksheet.Cells(10, 2) = " "
    objWorksheet.Cells(11, 2) = "Mobile Phase:"
    objWorksheet.Cells(12, 2) = "Wash Phase:"

'Data
    objWorksheet.Cells(14, 1) = "Data"
    objWorksheet.Cells(14, 1).Font.Bold = True
    objWorksheet.Cells(14, 3) = "Data A"
    objWorksheet.Cells(15, 3) = "(ug/s)"
    objWorksheet.Cells(14, 4) = "Data B"
    objWorksheet.Cells(15, 4) = "(ug/s)"
    objWorksheet.Cells(16, 2) = "1)"
    objWorksheet.Cells(17, 2) = "2)"
    objWorksheet.Cells(18, 2) = "3)"
    objWorksheet.Cells(19, 2) = "4)"
    objWorksheet.Cells(20, 2) = "5)"
    objWorksheet.Cells(21, 2) = "6)"
    objWorksheet.Cells(22, 2) = "7)"
    objWorksheet.Cells(23, 2) = "8)"
    objWorksheet.Cells(24, 2) = "9)"
    objWorksheet.Cells(25, 2) = "10)"

'Enter data
'Put your own data here, load a file, or something
'Data Set A
    objWorksheet.Cells(16, 4) = "111"
    objWorksheet.Cells(17, 4) = "222"
    objWorksheet.Cells(18, 4) = "333"
    objWorksheet.Cells(19, 4) = "444"
    objWorksheet.Cells(20, 4) = "555"
    objWorksheet.Cells(21, 4) = "666"
    objWorksheet.Cells(22, 4) = "777"
    objWorksheet.Cells(23, 4) = "888"
    objWorksheet.Cells(24, 4) = "999"
    objWorksheet.Cells(25, 4) = "1111"
'Data Set B
    objWorksheet.Cells(25, 3) = "111"
    objWorksheet.Cells(24, 3) = "222"
    objWorksheet.Cells(23, 3) = "333"
    objWorksheet.Cells(22, 3) = "444"
    objWorksheet.Cells(21, 3) = "555"
    objWorksheet.Cells(20, 3) = "666"
    objWorksheet.Cells(19, 3) = "777"
    objWorksheet.Cells(18, 3) = "888"
    objWorksheet.Cells(17, 3) = "999"
    objWorksheet.Cells(16, 3) = "1111"

'Draw Chart with the data
    Charts.Add
    ActiveChart.ChartType = xlLineMarkers
    ActiveChart.SetSourceData Source:=Sheets("Results").Range("A14:D25"), PlotBy _
        :=xlColumns
    ActiveChart.Location Where:=xlLocationAsObject, Name:="Results"
    With ActiveChart
        .HasTitle = True
        .ChartTitle.Characters.Text = objWorksheet.Cells(1, 1) ' Title
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Sample" ' X-Axis
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Amount " & objWorksheet.Cells(15, 3)  ' Y-Axis
    End With
    ActiveSheet.ChartObjects("Chart 1").Activate
    With ActiveSheet.Shapes("Chart 1")
        .Left = 240.75
        .Top = 178.5
    End With



'Turn back on alerts so user will be notified to save on exit.
'    objExcel.DisplayAlerts = True

'Free up memory, otherwise there will be a memory leak.
Set objExcel = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
End Sub

