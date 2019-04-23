---
title: ModelRelationships.Application property (Excel)
keywords: vbaxl10.chm939073
f1_keywords:
- vbaxl10.chm939073
ms.prod: excel
ms.assetid: 8c2d631a-84bc-8709-79ba-bffe40ed676f
ms.date: 04/20/2019
localization_priority: Normal
---


# ModelRelationships.Application property (Excel)

Returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[ModelRelationships](Excel.modelrelationships.md)** object.


## Property value

**APPLICATION**


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