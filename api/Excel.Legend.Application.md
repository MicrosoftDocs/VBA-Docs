---
title: Legend.Application property (Excel)
keywords: vbaxl10.chm621073
f1_keywords:
- vbaxl10.chm621073
ms.prod: excel
api_name:
- Excel.Legend.Application
ms.assetid: ced7136a-7f18-2024-d40c-61fdca244949
ms.date: 06/08/2017
localization_priority: Normal
---


# Legend.Application property (Excel)

When used without an object qualifier, this property returns an  **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

_expression_. `Application`

_expression_ A variable that represents a [Legend](Excel.Legend-graph-property.md) object.


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


[Legend Object](Excel.Legend(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]