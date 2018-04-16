'Semi-colong delineated'

Attribute VB_Name = "Oulook2Excel"
Option Explicit

Sub CopyToExcel()

  Dim xlApp As Object
  Dim xlWB As Object
  Dim xlSheet As Object
  Dim olItem As Outlook.MailItem
  Dim vText As Variant
  Dim sText As String
  Dim vItem As Variant
  Dim lRow As Long

  Dim i As Long
  Dim j As Long
  Dim rCount As Long
  Dim bXStarted As Boolean
  Const strPath As String = "[YOUR WORKBOOK PATH]" 'the path of the workbook

  If Application.ActiveExplorer.Selection.Count = 0 Then
    MsgBox "No Items selected!", vbCritical, "Error"
    Exit Sub
  End If
  On Error Resume Next
  Set xlApp = GetObject(, "Excel.Application")
  If Err <> 0 Then
    Application.StatusBar = "Please wait while Excel source is opened ... "
    Set xlApp = CreateObject("Excel.Application")
    bXStarted = True
  End If
  On Error GoTo 0

'Open the workbook to input the data
  Set xlWB = xlApp.Workbooks.Open(strPath)
  Set xlSheet = xlWB.Sheets("All Questions")

'Process each selected record
  For Each olItem In Application.ActiveExplorer.Selection
    sText = olItem.Body
    vText = Split(sText, "Semi-colon Delineated String for spreadsheet import") 'Get everything below this text in email'
    vItem = Split(Trim(vText(1)), Chr(59))
    
    'Find the next empty line of the worksheet
    lRow = xlSheet.Range("A1").CurrentRegion.Rows.Count

    For i = 0 To UBound(vItem)
      xlSheet.Cells(lRow + 1, i + 1) = Replace(vItem(i), vbLf, "")
    Next i
    
    'Resize cell to wrap text'
    For j = 1 To 9
    xlWB.Sheets("Q" & j).Range("B2:B100").WrapText = True
    Next j
   
    xlWB.Sheets("Additional Comments").Range("B2:B100").WrapText = True
    xlWB.Save
  Next olItem

  If bXStarted Then
      xlApp.Quit
  End If
  Set xlApp = Nothing
  Set xlWB = Nothing
  Set xlSheet = Nothing
  Set olItem = Nothing
End Sub
