---
title: ProtectedViewWindow.Workbook property (Excel)
keywords: vbaxl10.chm914084
f1_keywords:
- vbaxl10.chm914084
ms.prod: excel
api_name:
- Excel.ProtectedViewWindow.Workbook
ms.assetid: 379b98f0-b177-7910-4968-ce4ed2f1ca9d
ms.date: 05/09/2019
localization_priority: Normal
---


# ProtectedViewWindow.Workbook property (Excel)

Returns an object that represents the workbook that is open in the specified Protected View window. Read-only.


## Syntax

_expression_.**Workbook**

_expression_ A variable that represents a **[ProtectedViewWindow](Excel.ProtectedViewWindow.md)** object.


## Return value

**[Workbook](Excel.Workbook.md)**


## Remarks

Because a Protected View window is designed to protect the user from potentially malicious code, the operations that you can perform by using a **Workbook** object returned by the **Workbook** method will be limited. Any operation that is not allowed will return an error.

A workbook displayed in a Protected View window is not a member of the **[Workbooks](Excel.Workbooks.md)** collection. Instead, use the **Workbook** property to access a workbook that is displayed in a Protected View window.


## Example

The following example uses the **Workbook** property to return the workbook that is open in the first Protected View window.

```vb
Dim wbProtected As Workbook 
 
If Application.ProtectedViewWindows.Count > 0 Then 
    Set wbProtected = Application.ProtectedViewWindows(1).Workbook 
End If 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]