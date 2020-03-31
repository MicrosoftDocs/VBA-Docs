---
title: Calling the Windows API (differences in string function operations)
ROBOTS: INDEX
ms.prod: access
ms.assetid: ee882d00-46f5-2bfc-09fc-ce2941302c5e
ms.date: 06/08/2019
localization_priority: Normal
---


# Calling the Windows API (differences in string function operations)

**Applies to:** Access 2013 | Access 2016

The memory storage formats for text differ between Visual Basic for Applications (VBA) code and Access Basic code. (Access Basic was used in early versions of Microsoft Access.) Text is stored in ANSI format within Access Basic code and in Unicode format in Visual Basic. This topic discusses one potential issue when handling strings in the current version of Microsoft Access. For more information, see [Differences in String Function Operations](https://msdn.microsoft.com/library/40ce2b9a-cac6-589e-2b5e-d63be37efeee%28Office.15%29.aspx).

In several Windows API functions, the byte length of a string has a special meaning. For example, the following program returns a folder set up in Windows. In Microsoft Access, **LeftB** (Buffer, ret) does not return the correct string. This is because, in spite of the fact that it shows the byte length of an ANSI string, the **LeftB** function processes Unicode strings. In this case, use the **InStr** function so that only the character string, without nulls, is returned.

```vb
Private Declare Function GetWindowsDirectory Lib "kernel32" _ 
 Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _ 
 ByVal nSize As Long) As Long 
 
Private Sub Command1_Click() 
 Buffer$ = Space(255) 
 ret = GetWindowsDirectory(Buffer$, 255) 
 ' WinDir = LeftB(Buffer, ret) '<--- Incorrect code" 
 
 WinDir = Left(Buffer$, InStr(Buffer$, Chr(0)) - 1) 
 '<--Correct code" 
 Print WinDir 
End Sub
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]