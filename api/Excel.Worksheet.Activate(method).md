---
title: Worksheet.Activate method (Excel)
keywords: vbaxl10.chm174073
f1_keywords:
- vbaxl10.chm174073
ms.prod: excel
api_name:
- Excel.Worksheet.Activate
ms.assetid: b198dc36-99d0-42db-6cbb-7f68396fd2f5
ms.date: 06/08/2017
localization_priority: Priority
---


# Worksheet.Activate method (Excel)

Makes the current sheet the active sheet. 


## Syntax

_expression_. `Activate`

_expression_ A variable that represents a [Worksheet](./Excel.Worksheet.md) object.


## Remarks

Calling this method is equivalent to clicking the sheet's tab.


## Example

This example activates Sheet1.


```vb
Worksheets("Sheet1").Activate
```


## See also


[Worksheet Object](Excel.Worksheet.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]