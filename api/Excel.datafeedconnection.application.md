---
title: DataFeedConnection.Application property (Excel)
keywords: vbaxl10.chm927073
f1_keywords:
- vbaxl10.chm927073
ms.prod: excel
ms.assetid: 35fdc681-eb9e-cd3d-9e8f-712b5a6815f4
ms.date: 03/28/2019
localization_priority: Normal
---


# DataFeedConnection.Application property (Excel)

Returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[DataFeedConnection](Excel.datafeedconnection.md)** object.


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


## Property value

**APPLICATION**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]