---
title: PivotTable.TableRange1 property (Excel)
keywords: vbaxl10.chm235098
f1_keywords:
- vbaxl10.chm235098
ms.prod: excel
api_name:
- Excel.PivotTable.TableRange1
ms.assetid: 4dfea643-3299-82ee-a770-b961904eec7f
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.TableRange1 property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents the range containing the entire PivotTable report, but doesn't include page fields. Read-only.


## Syntax

_expression_.**TableRange1**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

The **[TableRange2](Excel.PivotTable.TableRange2.md)** property includes page fields.


## Example

This example selects all of the PivotTable report except its page fields.

```vb
Worksheets("Sheet1").Activate 
Range("A3").PivotTable.TableRange1.Select
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]