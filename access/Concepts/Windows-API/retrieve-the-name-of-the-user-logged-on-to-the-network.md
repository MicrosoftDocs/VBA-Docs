---
title: Retrieve the name of the user signed in to the network
ms.prod: access
ms.assetid: 3bf335a1-08d0-c8d5-8d89-36f0c29d47d0
ms.date: 09/26/2018
localization_priority: Normal
---


# Retrieve the name of the user signed in to the network

This topic contains a user-defined function, GetLogonName, that returns the current user name. The GetLogonName function utilizes the **GetUserNameA** Windows API to retrieve the current user name. 


```vb
' Access the GetUserNameA function in advapi32.dll and 
' call the function GetUserName. 
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _ 
 (ByVal lpBuffer As String, nSize As Long) As Long 
 
' Main routine to retrieve user name. 
Function GetLogonName() As String 
 
 ' Dimension variables 
 Dim lpBuff As String * 255 
 Dim ret As Long 
 
 ' Get the user name minus any trailing spaces found in the name. 
 ret = GetUserName(lpBuff, 255) 
 
 If ret > 0 Then 
 GetLogonName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1) 
 Else 
 GetLogonName = vbNullString 
 End If 
End Function
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
