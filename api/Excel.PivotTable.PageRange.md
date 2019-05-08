---
title: PivotTable.PageRange property (Excel)
keywords: vbaxl10.chm235087
f1_keywords:
- vbaxl10.chm235087
ms.prod: excel
api_name:
- Excel.PivotTable.PageRange
ms.assetid: 05629703-c43f-282c-e4da-22c95094e15b
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.PageRange property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents the range that contains the page area in the PivotTable report. Read-only.


## Syntax

_expression_.**PageRange**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Example

This example selects the page headers in the PivotTable report.

```vb
Worksheets("Sheet1").Activate 
Range("A3").Select 
ActiveCell.PivotTable.PageRange.Select
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]