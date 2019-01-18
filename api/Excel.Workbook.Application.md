---
title: Workbook.Application property (Excel)
keywords: vbaxl10.chm198073
f1_keywords:
- vbaxl10.chm198073
ms.prod: excel
api_name:
- Excel.Workbook.Application
ms.assetid: 91b30f9d-48e5-e033-8daf-416d1c0e547d
ms.date: 06/08/2017
localization_priority: Normal
---


# Workbook.Application property (Excel)

When used without an object qualifier, this property returns an  **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

_expression_. `Application`

_expression_ A variable that represents a [Workbook](./Excel.Workbook.md) object.


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


## See also


[Workbook Object](Excel.Workbook.md)

