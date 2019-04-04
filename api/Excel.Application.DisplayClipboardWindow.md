---
title: Application.DisplayClipboardWindow property (Excel)
keywords: vbaxl10.chm133093
f1_keywords:
- vbaxl10.chm133093
ms.prod: excel
api_name:
- Excel.Application.DisplayClipboardWindow
ms.assetid: 16686caf-39ed-90fa-4a61-92b3f825cc6c
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.DisplayClipboardWindow property (Excel)

Returns **True** if the Microsoft Office Clipboard can be displayed. Read/write **Boolean**.


## Syntax

_expression_.**DisplayClipboardWindow**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

In this example, Microsoft Excel determines if the Office Clipboard can be displayed, and notifies the user.

```vb
Sub SeeClipboard() 
 
 ' Determine if Office Clipboard can be displayed. 
 If Application.DisplayClipboardWindow = True Then 
 MsgBox "Office Clipboard can be displayed." 
 Else 
 MsgBox "Office Clipboard cannot be displayed." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]