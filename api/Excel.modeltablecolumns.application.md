---
title: ModelTableColumns.Application property (Excel)
keywords: vbaxl10.chm931073
f1_keywords:
- vbaxl10.chm931073
ms.prod: excel
ms.assetid: cb086ea8-fcce-8c36-a92c-d006b774ff82
ms.date: 05/01/2019
localization_priority: Normal
---


# ModelTableColumns.Application property (Excel)

Returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[ModelTableColumns](Excel.modeltablecolumns.md)** object.


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


## Property value

**APPLICATION**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]