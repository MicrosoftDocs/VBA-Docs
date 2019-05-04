---
title: PivotCell.Application property (Excel)
keywords: vbaxl10.chm691073
f1_keywords:
- vbaxl10.chm691073
ms.prod: excel
api_name:
- Excel.PivotCell.Application
ms.assetid: 63a2a460-a72f-086a-8401-2e029847ee39
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotCell.Application property (Excel)

When used without an object qualifier, this property returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. 

When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[PivotCell](Excel.PivotCell.md)** object.


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