---
title: Workbook.CanCheckIn method (Excel)
keywords: vbaxl10.chm199205
f1_keywords:
- vbaxl10.chm199205
ms.prod: excel
api_name:
- Excel.Workbook.CanCheckIn
ms.assetid: 17f7cbdd-0ce0-8e3a-46f3-cb6dafaaa40a
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.CanCheckIn method (Excel)

**True** if Microsoft Excel can check in a specified workbook to a server. Read/write **Boolean**.


## Syntax

_expression_.**CanCheckIn**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Return value

**Boolean**


## Example

This example checks the server to see if the specified workbook can be checked in. If it can be, it saves and closes the workbook and checks it back into the server.

```vb
Sub CheckInOut(strWkbCheckIn As String) 
 
 ' Determine if workbook can be checked in. 
 If Workbooks(strWkbCheckIn).CanCheckIn = True Then 
 Workbooks(strWkbCheckIn).CheckIn 
 MsgBox strWkbCheckIn & " has been checked in." 
 Else 
 MsgBox "This file cannot be checked in " & _ 
 "at this time. Please try again later." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]