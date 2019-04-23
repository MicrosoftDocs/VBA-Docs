---
title: Characters.Application property (Excel)
keywords: vbaxl10.chm251073
f1_keywords:
- vbaxl10.chm251073
ms.prod: excel
api_name:
- Excel.Characters.Application
ms.assetid: b67b835c-9324-04ed-4d7f-e1756e6fe7f8
ms.date: 04/16/2019
localization_priority: Normal
---


# Characters.Application property (Excel)

When used without an object qualifier, this property returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. 

When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[Characters](Excel.Characters.md)** object.


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