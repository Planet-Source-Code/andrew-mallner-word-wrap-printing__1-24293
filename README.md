<div align="center">

## Word Wrap Printing


</div>

### Description

This is a subroutine to automatically wrap a text string in whole words to a printer.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Andrew Mallner](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andrew-mallner.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andrew-mallner-word-wrap-printing__1-24293/archive/master.zip)





### Source Code

```
<pre>Private Sub PrintWordWrap(YourText As String, LeftMargin_InTwips As Long, RightMargin_InTwips As Long)
On Error GoTo Errors
Start = 1
Char = ""
TempText = ""
Dim boolSpace As Boolean
For Location = 1 To Len(YourText)
Char = Mid(YourText, Location, 1)
If Char = " " Then
 If Printer.TextWidth(TempText2 & Mid(YourText, Start, Location - Start)) <= Printer.Width - RightMargin_InTwips - LeftMargin_InTwips - 700 Then
  TempText = Mid(YourText, Start, Location - Start)
  Pos = Location
  boolSpace = True
 Else
  Start = Location
  Pos2 = Location
  Printer.CurrentX = LeftMargin_InTwips
  Printer.Print TempText2 & TempText
  TempText2 = Mid(YourText, Pos + 1, Location - Pos - 1)
  TempText = ""
  boolSpace = False
 End If
ElseIf Char = vbCr And Mid(YourText, Location + 1, 1) = vbLf And Printer.TextWidth(TempText2 & Mid(YourText, Start, Location - Start)) <= Printer.Width - RightMargin_InTwips - LeftMargin_InTwips - 700 Then
 If Not InStr(Mid(YourText, Start, Location - Start), vbCr) <> 0 Then
 Printer.CurrentX = LeftMargin_InTwips
 End If
 Printer.Print TempText2 & Mid(YourText, Start, Location - Start);
 Start = Location + 1
 Pos2 = Location
 TempText = ""
 TempText2 = ""
 boolSpace = False
ElseIf boolSpace = False And _
  Printer.TextWidth(Mid(YourText, Start, Location - Start)) >= Printer.Width - Printer.TextWidth("W") - RightMargin_InTwips - LeftMargin_InTwips - 700 And _
  Printer.TextWidth(Mid(YourText, Start, Location - Start)) < Printer.Width - RightMargin_InTwips - LeftMargin_InTwips - 700 Then
 Printer.CurrentX = LeftMargin_InTwips
 Printer.Print Mid(YourText, Start, Location - Start)
 Start = Location
 Pos = Location
 TempText = ""
 TempText2 = ""
End If
If Printer.CurrentY > Printer.Height Then Printer.NewPage
Next
If Printer.TextWidth(TempText2 & TempText) <= Printer.Width - RightMargin_InTwips - LeftMargin_InTwips - 700 Then
 Printer.CurrentX = LeftMargin_InTwips
 Printer.Print TempText2 & Mid(YourText, Pos2, Location - Pos2);
End If
Printer.EndDoc
Exit Sub
Errors:
boxit = MsgBox(Err.Description, vbOKOnly + vbApplicationModal + vbInformation, Err.Source & " Error #" & Err.Number)
' 700 twips are subtracted from the width of the
' page to account for the non-printable area for
' MY printer. I don't know for sure, but this may
' vary depending on your printer.
End Sub</pre>
```

