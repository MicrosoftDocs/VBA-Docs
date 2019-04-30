---
title: Model.Application property (Excel)
keywords: vbaxl10.chm941073
f1_keywords:
- vbaxl10.chm941073
ms.prod: excel
ms.assetid: 201e79d9-8e9b-0ff4-5f6e-dcdd3911b08a
ms.date: 04/30/2019
localization_priority: Normal
---


# Model.Application property (Excel)

Returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[Model](Excel.Model.md)** object.


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