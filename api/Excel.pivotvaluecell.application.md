---
title: PivotValueCell.Application property (Excel)
keywords: vbaxl10.chm917073
f1_keywords:
- vbaxl10.chm917073
ms.assetid: f749fa87-4b7f-4609-13dd-190888da6233
ms.date: 05/07/2019
ms.localizationpriority: medium
---


# PivotValueCell.Application property (Excel)

Returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[PivotValueCell](Excel.pivotvaluecell.md)** object.


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