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
