---
title: ChartObject.Application Property (Excel)
keywords: vbaxl10.chm493073
f1_keywords:
- vbaxl10.chm493073
ms.prod: excel
api_name:
- Excel.ChartObject.Application
ms.assetid: 1d2e8fde-3641-fd19-bf7d-f4b72d827e22
ms.date: 06/08/2017
---


# ChartObject.Application Property (Excel)

When used without an object qualifier, this property returns an  **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _expression_. `Application`

 _expression_ A variable that represents a [ChartObject](Excel.ChartObject.md) object.


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


[ChartObject Object](Excel.ChartObject.md)

