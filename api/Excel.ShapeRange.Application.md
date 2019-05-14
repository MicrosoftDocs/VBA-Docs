---
title: ShapeRange.Application property (Excel)
keywords: vbaxl10.chm639073
f1_keywords:
- vbaxl10.chm639073
ms.prod: excel
api_name:
- Excel.ShapeRange.Application
ms.assetid: 1fd7d168-7d4e-c732-db7d-986757a0a483
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.Application property (Excel)

When used without an object qualifier, this property returns an  **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[ShapeRange](Excel.shaperange.md)** object.


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


## See also


[ShapeRange Object](Excel.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]