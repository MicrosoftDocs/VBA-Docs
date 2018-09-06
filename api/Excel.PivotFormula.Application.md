---
title: PivotFormula.Application Property (Excel)
keywords: vbaxl10.chm230073
f1_keywords:
- vbaxl10.chm230073
ms.prod: excel
api_name:
- Excel.PivotFormula.Application
ms.assetid: 5733d267-64ef-5891-7487-fd08e3d3c2c3
ms.date: 06/08/2017
---


# PivotFormula.Application Property (Excel)

When used without an object qualifier, this property returns an  **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _expression_. `Application`

 _expression_ A variable that represents a [PivotFormula](Excel.PivotFormula.md) object.


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


[PivotFormula Object](Excel.PivotFormula.md)

