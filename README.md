<div align="center">

## A Split Procedure


</div>

### Description

Splits a string into an array. If you send a " " it will split all the words into each array position.
 
### More Info
 
The string to split.

The splitter, e.g. " "

Private Sub Command1_Click()

Dim SplitReturn As Variant

SplitReturn = Splitter(Text1.Text, " ")

MsgBox SplitReturn(1)

End Sub

Returns an array of the results.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Paul Spiteri](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/paul-spiteri.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/paul-spiteri-a-split-procedure__1-9165/archive/master.zip)





### Source Code

```
Public Function Splitter(SplitString As String, SplitLetter As String) As Variant
 ReDim SplitArray(1 To 1) As Variant
 Dim TempLetter As String
 Dim TempSplit As String
 Dim i As Integer
 Dim x As Integer
 Dim StartPos As Integer
 SplitString = SplitString & SplitLetter
 For i = 1 To Len(SplitString)
  TempLetter = Mid(SplitString, i, Len(SplitLetter))
  If TempLetter = SplitLetter Then
   TempSplit = Mid(SplitString, (StartPos + 1), (i - StartPos) - 1)
   If TempSplit <> "" Then
    x = x + 1
    ReDim Preserve SplitArray(1 To x) As Variant
    SplitArray(x) = TempSplit
   End If
   StartPos = i
  End If
 Next i
 Splitter = SplitArray
End Function
```

