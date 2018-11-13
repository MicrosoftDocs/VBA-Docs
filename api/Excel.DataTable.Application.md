---
title: DataTable.Application property (Excel)
keywords: vbaxl10.chm625073
f1_keywords:
- vbaxl10.chm625073
ms.prod: excel
api_name:
- Excel.DataTable.Application
ms.assetid: 0bc376d3-bf52-5b23-c737-60dc5b3bf899
ms.date: 06/08/2017
---


# DataTable.Application property (Excel)

When used without an object qualifier, this property returns an  **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _expression_. `Application`

 _expression_ A variable that represents a [DataTable](Excel.DataTable-graph-property.md) object.


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


[DataTable Object](Excel.DataTable(object).md)

