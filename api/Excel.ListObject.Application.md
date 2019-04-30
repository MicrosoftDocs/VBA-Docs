---
title: ListObject.Application property (Excel)
keywords: vbaxl10.chm733073
f1_keywords:
- vbaxl10.chm733073
ms.prod: excel
api_name:
- Excel.ListObject.Application
ms.assetid: 72c071cf-e0c8-fcc5-5207-93ca43272dfa
ms.date: 04/30/2019
localization_priority: Normal
---


# ListObject.Application property (Excel)

When used without an object qualifier, this property returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. 

When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[ListObject](Excel.ListObject.md)** object.


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