---
title: TimelineViewState.Application property (Excel)
keywords: vbaxl10.chm951073
f1_keywords:
- vbaxl10.chm951073
ms.prod: excel
ms.assetid: b00518cc-b584-c562-0ae3-cf1e24844bdd
ms.date: 05/18/2019
localization_priority: Normal
---


# TimelineViewState.Application property (Excel)

Returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[TimelineViewState](Excel.TimelineViewState.md)** object.


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