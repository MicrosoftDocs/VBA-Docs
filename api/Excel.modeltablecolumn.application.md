---
title: ModelTableColumn.Application property (Excel)
keywords: vbaxl10.chm929073
f1_keywords:
- vbaxl10.chm929073
ms.prod: excel
ms.assetid: 69540e35-6a9a-0fd9-23b1-31457b33ba68
ms.date: 05/01/2019
localization_priority: Normal
---


# ModelTableColumn.Application property (Excel)

Returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[ModelTableColumn](Excel.modeltablecolumn.md)** object.


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