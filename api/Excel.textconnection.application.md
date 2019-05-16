---
title: TextConnection.Application property (Excel)
keywords: vbaxl10.chm925073
f1_keywords:
- vbaxl10.chm925073
ms.prod: excel
ms.assetid: a3dc9071-4d42-6293-b9df-25dcc84d4ca8
ms.date: 05/17/2019
localization_priority: Normal
---


# TextConnection.Application property (Excel)

Returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[TextConnection](Excel.TextConnection.md)** object.


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