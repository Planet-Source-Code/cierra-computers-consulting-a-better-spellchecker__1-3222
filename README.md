<div align="center">

## A Better SpellChecker


</div>

### Description

This is basically an enhanced version of the SpellCheck function that I found in MSDN from Microsoft. They left out a couple things.
 
### More Info
 
Text

Be sure to Reference the MS Word Object Library


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Cierra Computers & Consulting](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/cierra-computers-consulting.md)
**Level**          |Unknown
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/cierra-computers-consulting-a-better-spellchecker__1-3222/archive/master.zip)





### Source Code

```
Public Function SpellCheck(strText As String, Optional blnSupressMsg As Boolean = False) As String
'This function opens the MS Word Object and uses its spell checker
'passing back the corrected string
On Error Resume Next
Dim oWDBasic As Object
Dim sTmpString As String
If strText = "" Then
   If blnSupressMsg = False Then
     MsgBox "Nothing to spell check.", vbInformation, App.ProductName
   End If
   Exit Function
End If
Screen.MousePointer = vbHourglass
Set oWDBasic = CreateObject("Word.Basic")
With oWDBasic
   .FileNew
   .Insert strText
   .ToolsSpelling oWDBasic.EditSelectAll
   .SetDocumentVar "MyVar", oWDBasic.Selection
End With
sTmpString = oWDBasic.GetDocumentVar("MyVar")
sTmpString = Left(sTmpString, Len(sTmpString) - 1)
If sTmpString = "" Then
   SpellCheck = strText
Else
   SpellCheck = sTmpString
End If
oWDBasic.FileCloseAll 2
oWDBasic.AppClose
Set oWDBasic = Nothing
Screen.MousePointer = vbNormal
If blnSupressMsg = False Then
   MsgBox "Spell check is completed.", vbInformation, App.ProductName
End If
End Function
```

