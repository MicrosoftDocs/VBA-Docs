---
title: ModelTableColumns.Application property (Excel)
keywords: vbaxl10.chm931073
f1_keywords:
- vbaxl10.chm931073
ms.prod: excel
ms.assetid: cb086ea8-fcce-8c36-a92c-d006b774ff82
ms.date: 06/08/2017
localization_priority: Normal
---


# ModelTableColumns.Application property (Excel)

Returns an  **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_. `Application`

_expression_ A variable that represents a [ModelTableColumns object (Excel)](Excel.modeltablecolumns.md) object.


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



[ModelTableColumns Object](Excel.modeltablecolumns.md)

