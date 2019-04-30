---
title: ModelColumnChange.Application property (Excel)
keywords: vbaxl10.chm965073
f1_keywords:
- vbaxl10.chm965073
ms.prod: excel
ms.assetid: 42065d25-aaef-e92a-f174-47f056e1e460
ms.date: 05/01/2019
localization_priority: Normal
---


# ModelColumnChange.Application property (Excel)

Returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[ModelColumnChange](Excel.modelcolumnchange.md)** object.


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