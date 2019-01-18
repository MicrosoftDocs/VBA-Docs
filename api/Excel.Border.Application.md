---
title: Border.Application property (Excel)
keywords: vbaxl10.chm546073
f1_keywords:
- vbaxl10.chm546073
ms.prod: excel
api_name:
- Excel.Border.Application
ms.assetid: 61d227a1-e6da-4fc0-5bf3-ca815c1c8d44
ms.date: 06/08/2017
localization_priority: Normal
---


# Border.Application property (Excel)

When used without an object qualifier, this property returns an  **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

_expression_. `Application`

_expression_ A variable that represents a [Border](Excel.Border-graph-property.md) object.


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


[Border Object](Excel.Border(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]