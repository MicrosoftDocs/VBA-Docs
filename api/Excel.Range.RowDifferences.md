---
title: Range.RowDifferences method (Excel)
keywords: vbaxl10.chm144189
f1_keywords:
- vbaxl10.chm144189
ms.prod: excel
api_name:
- Excel.Range.RowDifferences
ms.assetid: 89030ca3-9f59-7426-d050-89dcabf00887
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.RowDifferences method (Excel)

Returns a **Range** object that represents all the cells whose contents are different from those of the comparison cell in each row.


## Syntax

_expression_.**RowDifferences** (_Comparison_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Comparison_|Required| **Variant**|A single cell to compare with the specified range.|

## Return value

Range


## Example

This example selects the cells in row one on Sheet1 whose contents are different from those of cell D1.

```vb
Worksheets("Sheet1").Activate 
Set c1 = ActiveSheet.Rows(1).RowDifferences( _ 
 comparison:=ActiveSheet.Range("D1")) 
c1.Select
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]