---
title: ModelRelationship.Application property (Excel)
keywords: vbaxl10.chm937073
f1_keywords:
- vbaxl10.chm937073
ms.prod: excel
ms.assetid: fc6832ad-4100-e1ac-f286-6f0cbe11c983
ms.date: 04/20/2019
localization_priority: Normal
---


# ModelRelationship.Application property (Excel)

Returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[ModelRelationship](Excel.modelrelationship.md)** object.


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