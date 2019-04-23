---
title: TableObject.Application property (Excel)
keywords: vbaxl10.chm915073
f1_keywords:
- vbaxl10.chm915073
ms.prod: excel
ms.assetid: 7150f52d-c871-12bc-89d8-42993844187d
ms.date: 04/19/2019
localization_priority: Normal
---


# TableObject.Application property (Excel)

Returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[TableObject](Excel.tableobject.md)** object.


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