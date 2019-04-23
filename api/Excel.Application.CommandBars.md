---
title: Application.CommandBars property (Excel)
keywords: vbaxl10.chm183088
f1_keywords:
- vbaxl10.chm183088
ms.prod: excel
api_name:
- Excel.Application.CommandBars
ms.assetid: b1884d43-557b-47be-1cef-20404069b576
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.CommandBars property (Excel)

Returns a **[CommandBars](Office.CommandBars.md)** object that represents the Microsoft Excel command bars. Read-only.


## Syntax

_expression_.**CommandBars**

_expression_ An expression that returns an **[Application](Excel.Application(object).md)** object.


## Remarks

Used with the **Application** object, this property returns the set of built-in and custom command bars available to the application.

When a workbook is embedded in another application and activated by the user by double-clicking the workbook, using this property with a **[Workbook](Excel.Workbook.md)** object returns the set of Microsoft Excel command bars available within the other application. At all other times, using this property with a **Workbook** object returns **Nothing**.

There is no programmatic way to return the set of command bars attached to a workbook.


## Example

This example deletes all custom command bars that aren't visible.

```vb
For Each bar In Application.CommandBars 
    If Not bar.BuiltIn And Not bar.Visible Then bar.Delete 
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
