<div align="center">

## INI Reading and Writing Made Simple\!


</div>

### Description

This code was designed for reading and writing INI files. You put it in a module called modINI.
 
### More Info
 
Syntax:

ReadINI("SECTION", "FIELD", filename)

WriteINI("SECTION", "FIELD", "VALUE", filename)

This is very "newbie" friendly code. Any user, from beginner to advanced user, can use it.

The value of the field read from the INI for ReadINI. For WriteINI, it will return nothing.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Austen Frazier](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/austen-frazier.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/austen-frazier-ini-reading-and-writing-made-simple__1-61508/archive/master.zip)





### Source Code

```
Option Explicit
'ModINI.Bas
'INI reading/writing
Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal FileName$)
Public Sub WriteINI(INISection As String, INIKey As String, INIValue As String, INIFile As String)
  Call WritePrivateProfileString(INISection, INIKey, INIValue, INIFile)
End Sub
Public Function ReadINI(INISection As String, INIKey As String, INIFile As String) As String
  Dim StringBuffer As String
  Dim StringBufferSize As Long
  StringBuffer = Space$(255)
  StringBufferSize = Len(StringBuffer)
  StringBufferSize = GetPrivateProfileString(INISection, INIKey, "", StringBuffer, StringBufferSize, INIFile)
  If StringBufferSize > 0 Then
    ReadINI = Left$(StringBuffer, StringBufferSize)
  Else
    ReadINI = ""
  End If
End Function
```

