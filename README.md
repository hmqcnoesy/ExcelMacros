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
This will use Beyond Compare (assuming install at `C:\Program Files\Beyond Compare 4\BCompare.exe`) to compare the currently active sheet with the next sheet of the active workbook.  Make sure the temp file location specified in the bottom of the code already exists.

```vba
Public Sub CompareWorksheets()
    Dim ws As Worksheet
    Dim index As Integer
    Dim path As String
    Dim path1 As String
    Dim path2 As String
    Dim sh As Variant
    
    Set ws = ActiveSheet
    index = ws.index
    
    If ActiveWorkbook.Sheets.Count < (index + 1) Then
        MsgBox "Active workbook doesn't have sheet after selected sheet."
        Exit Sub
    End If
        
    path = "C:\temp"
    
    Set ws = ActiveWorkbook.Sheets(index)
    path1 = path & "\" & ws.Name & ".xlsx"
    ws.Copy
    Application.DisplayAlerts = False
    With ActiveWorkbook
        .SaveAs Filename:=path1, FileFormat:=xlOpenXMLWorkbook
        .Close SaveChanges:=False
    End With
    
    Set ws = ActiveWorkbook.Sheets(index + 1)
   path2 = path & "\" & ws.Name & ".xlsx"
    ws.Copy
    With ActiveWorkbook
        .SaveAs Filename:=path2, FileFormat:=xlOpenXMLWorkbook
        .Close SaveChanges:=False
    End With
    Application.DisplayAlerts = True
    
    sh = Shell("""C:\Program Files\Beyond Compare 4\BCompare.exe"" """ & path1 & """ """ & path2 & """", vbNormalFocus)
    
End Sub

Public Sub CompareWorksheetsAsText()
    Dim ws As Worksheet
    Dim index As Integer
    Dim path As String
    Dim path1 As String
    Dim path2 As String
    Dim sh As Variant
    
    Set ws = ActiveSheet
    index = ws.index
    
    If ActiveWorkbook.Sheets.Count < (index + 1) Then
        MsgBox "Active workbook doesn't have sheet after selected sheet."
        Exit Sub
    End If
        
    path = "C:\temp"
    
    Set ws = ActiveWorkbook.Sheets(index)
    path1 = path & "\" & ws.Name & ".txt"
    ws.Copy
    Application.DisplayAlerts = False
    With ActiveWorkbook
        .SaveAs Filename:=path1, FileFormat:=xlText
        .Close SaveChanges:=False
    End With
    
    Set ws = ActiveWorkbook.Sheets(index + 1)
    path2 = path & "\" & ws.Name & ".txt"
    ws.Copy
    With ActiveWorkbook
        .SaveAs Filename:=path2, FileFormat:=xlText
        .Close SaveChanges:=False
    End With
    Application.DisplayAlerts = True
    
    sh = Shell("""C:\Program Files\Beyond Compare 4\BCompare.exe"" """ & path1 & """ """ & path2 & """", vbNormalFocus)
    
End Sub
```


## Summarize Timecard
Summarizes timecard in my own personal format.
Add reference to `Microsoft.Scripting.Runtime`.

```vba
Sub SummarizeTimecard()
    Dim dates As Scripting.Dictionary
    Dim pos As Scripting.Dictionary
    Dim dateKey As Variant
    Dim poKey As Variant
    Dim row As Integer
    Dim message As String
        
    row = 3
    Set dates = New Scripting.Dictionary
    
    Do Until Cells(row, 2).Value = ""
        If ActiveSheet.Cells(row, 7) <> "" Then
            If ActiveSheet.Cells(row, 5) <> "" Then
                poKey = ActiveSheet.Cells(row, 5).Value
            Else
                poKey = ActiveSheet.Cells(row, 3).Value
            End If
                
            If ActiveSheet.Cells(row, 1).Value <> "" Then
                Set pos = New Scripting.Dictionary
                pos.Add poKey, ActiveSheet.Cells(row, 7)
                dates.Add ActiveSheet.Cells(row, 1).Value, pos
                lastDate = ActiveSheet.Cells(row, 1).Value
            Else
                If dates(lastDate).Exists(poKey) Then
                    dates(lastDate)(poKey) = dates(lastDate)(poKey) + ActiveSheet.Cells(row, 7)
                Else
                    Set pos = New Scripting.Dictionary
                    dates(lastDate).Add poKey, ActiveSheet.Cells(row, 7)
                End If
            End If
        
        End If
        row = row + 1
        
    Loop
    
    For Each dateKey In dates.Keys
        message = message & dateKey & Chr(13) & Chr(10)
        For Each poKey In dates(dateKey)
            message = message & Chr(9) & poKey & " :  " & Round(dates(dateKey)(poKey), 2) & Chr(13) & Chr(10)
        Next
    Next
    
    MsgBox message
    
    Set dates = Nothing
    Set pos = Nothing
End Sub
```
