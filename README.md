# ExcelMacros
Collection of useful Excel macros and snippets

## Format IIS Logs
Drop an IIS Log into Excel, then this macro will convert the text to columns and clean up the formatting

```vba
Sub FormatIISLog()
  Dim i As Integer
  Rows("1:3").Delete Shift:=xlUp
  Range("A1").Value = Replace(Range("A1").Value, "#Fields: ", "")
  Columns("A:A").TextToColumns _
    Destination:=Range("A1"), _
    DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, _
    ConsecutiveDelimiter:=False, _
    Tab:=False, Semicolon:=False, Comma:=False, Space:=True, Other:=False
    
  i = 1
  While (Range("A" & i).Value <> "")
    If Left(Range("A" & i), 1) = "#" Then
      Rows(i & ":" & i).Delete Shift:=xlUp
    Else
      i = i + 1
    End If
  Wend
  Range("A:A,B:B,C:C,D:D,E:E,H:H,I:I,L:L,M:M,N:N,O:O").EntireColumn.AutoFit
End Sub
```


## Compare Worksheets
This will use Beyond Compare (assuming install at `c:\apps\Beyond Compare 4\BCompare.exe`) to compare sheets 1 and 2 of the active workbook.

```vba
Public Sub CompareWorksheets()
  Dim ws As Worksheet    Dim path As String    Dim path1 As String    Dim path2 As String    Dim sh As Variant        If ActiveWorkbook.Sheets.Count < 2 Then        MsgBox "Active workbook doesn't have 2 sheets."        Exit Sub    End If        path = "C:\temp"        Set ws = ActiveWorkbook.Sheets(1)    path1 = path & "\" & ws.Name & ".xlsx"    ws.Copy    Application.DisplayAlerts = False    With ActiveWorkbook        .SaveAs Filename:=path1, FileFormat:=xlOpenXMLWorkbook        .Close SaveChanges:=False    End With        Set ws = ActiveWorkbook.Sheets(2)    path2 = path & "\" & ws.Name & ".xlsx"    ws.Copy    With ActiveWorkbook        .SaveAs Filename:=path2, FileFormat:=xlOpenXMLWorkbook        .Close SaveChanges:=False    End With    Application.DisplayAlerts = True        sh = Shell("""C:\apps\Beyond Compare 4\BCompare.exe"" """ & path1 & """ """ & path2 & """", vbNormalFocus)    End Sub
```
