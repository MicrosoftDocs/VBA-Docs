---
title: Application.hWnd property (Excel)
keywords: vbaxl10.chm133277
f1_keywords:
- vbaxl10.chm133277
api_name:
- Excel.Application.Hwnd
ms.assetid: ed98b59c-1ebf-f319-f986-3406e4fdb766
ms.date: 04/04/2019
ms.localizationpriority: medium
---


# Application.hWnd property (Excel)

Returns a **Long** indicating the top-level window handle of the Microsoft Excel window. Read-only.


## Syntax

_expression_.**hWnd**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

In this example, Microsoft Excel notifies the user of the top-level window handle of the Excel window.

```vb
Sub CheckHwnd() 
 
 MsgBox "The top-level window handle is: " & _ 
 Application.hWnd 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
