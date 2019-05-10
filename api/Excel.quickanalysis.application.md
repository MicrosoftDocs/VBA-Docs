---
title: QuickAnalysis.Application property (Excel)
keywords: vbaxl10.chm919073
f1_keywords:
- vbaxl10.chm919073
ms.prod: excel
ms.assetid: ad51f454-62a0-7eb7-b629-b72bd000e0e9
ms.date: 05/10/2019
localization_priority: Normal
---


# QuickAnalysis.Application property (Excel)

Returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[QuickAnalysis](Excel.quickanalysis.md)** object.


## Property value

**APPLICATION**


## Example

This example displays a message about the application that created _myObject_.

```vb
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]