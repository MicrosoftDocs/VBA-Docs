---
title: PivotValueCell.Application property (Excel)
keywords: vbaxl10.chm917073
f1_keywords:
- vbaxl10.chm917073
ms.prod: excel
ms.assetid: f749fa87-4b7f-4609-13dd-190888da6233
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotValueCell.Application property (Excel)

Returns an  **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_. `Application`

_expression_ A variable that represents a [PivotValueCell object (Excel)](Excel.pivotvaluecell.md) object.


## Example

This example displays a message about the application that created  `myObject`.


```vb
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If
```


## Property value

 **APPLICATION**


## See also



[PivotValueCell Object](Excel.pivotvaluecell.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]