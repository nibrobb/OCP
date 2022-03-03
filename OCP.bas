Attribute VB_Name = "OCP"
Sub LoopAllExcelFilesInFolder()
'PURPOSE: To loop through all Excel files in a user specified folder and perform a set task on them
'SOURCE: www.TheSpreadsheetGuru.com

Dim masterWS As Worksheet

Dim wb As Workbook
Dim myPath As String
Dim myFile As String
Dim myExtension As String
Dim FldrPicker As FileDialog
Dim ukeNr As Integer

'Define masterWS to be the initial worksheet from which we started the macro
Set masterWS = ActiveSheet

'Start på uke
ukeNr = 1

'Filnavn som skal formateres
Dim formattedFileName As String


'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual

'Retrieve Target Folder Path From User
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xlsx"

'Target Path with Ending Extention
  myFile = Dir(myPath & myExtension)

'Put in a little discriptive header
  masterWS.Range("A1").Value = "Uke"
  masterWS.Range("B1").Value = "Brutto lønn"
  masterWS.Range("C1").Value = "Netto lønn"
  masterWS.Range("D1").Value = "Innkjørt"

'Loop through each Excel file in folder
  Do While myFile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(Filename:=myPath & myFile)
    
    'Format filename for use in column A
    formattedFileName = Replace(myFile, "Uke ", "")
    formattedFileName = Replace(formattedFileName, ".xlsx", "")
    formattedFileName = Replace(formattedFileName, " - del 1", "")
    formattedFileName = Replace(formattedFileName, " - del 2", "")
    formattedFileName = Replace(formattedFileName, " - Robin", "")
    
    'Ensure Workbook has opened before moving on to next line of code
      DoEvents
    
    
    
    'Value of cells in column is equal to filename
    
    'The "master-workbook" should be ThisWorkbook, right?
    'And the sheet to pase things in, I guess we'll just call it DATA
    
    masterWS.Range("A" & CStr(ukeNr + 1)).Value = formattedFileName
    
    ' Replace ( string1, find, replacement, [start, [count, [compare]]] )
    
    'Copy/paste brutto lønn + tips (cell K24)
    wb.Worksheets("uke1").Range("K24").Copy
    masterWS.Range("B" & CStr(ukeNr + 1)).PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats
    
    'Copy/paste netto lønn (cell K26)
    wb.Worksheets("uke1").Range("K26").Copy
    masterWS.Range("C" & CStr(ukeNr + 1)).PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats
    
    'Copy/paste brutto innkjørt (cell G17)
    wb.Worksheets("uke1").Range("G17").Copy
    masterWS.Range("D" & CStr(ukeNr + 1)).PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats
    
    'Save and Close Workbook
      wb.Close SaveChanges:=False
      
    'Ensure Workbook has closed before moving on to next line of code
      DoEvents

    'Get next file name
      myFile = Dir
      
    'Increment ukeNr by one
      ukeNr = ukeNr + 1
  Loop

'Message Box when tasks are completed
  MsgBox "Task Complete!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
