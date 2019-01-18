---
title: Range.Calculate method (Excel)
keywords: vbaxl10.chm144090
f1_keywords:
- vbaxl10.chm144090
ms.prod: excel
api_name:
- Excel.Range.Calculate
ms.assetid: 7c29afda-4980-6992-fc8d-b4caf2f74660
ms.date: 06/08/2017
localization_priority: Priority
---


# Range.Calculate method (Excel)

Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet, as shown in the following table.


## Syntax

_expression_. `Calculate`

_expression_ A variable that represents a [Range](excel.range-graph-property.md) object.


## Return value

Variant


## Remarks





|**To calculate**|**Follow this example**|
|:-----|:-----|
|All open workbooks| `Application.Calculate` (or just `Calculate`)|
|A specific worksheet| `Worksheets(1).Calculate`|
|A specified range| `Worksheets(1).Rows(2).Calculate`|

## Example

This example calculates the formulas in columns A, B, and C in the used range on Sheet1.


```vb
Worksheets("Sheet1").UsedRange.Columns("A:C").Calculate
```


## See also


[Range Object](Excel.Range(object).md)

