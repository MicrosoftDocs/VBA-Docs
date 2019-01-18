---
title: Range.Merge method (Excel)
keywords: vbaxl10.chm144158
f1_keywords:
- vbaxl10.chm144158
ms.prod: excel
api_name:
- Excel.Range.Merge
ms.assetid: eff315d8-fa8f-e452-2bcd-15be4d97a077
ms.date: 06/08/2017
localization_priority: Priority
---


# Range.Merge method (Excel)

Creates a merged cell from the specified  **[Range](Excel.Range(object).md)** object.


## Syntax

_expression_. `Merge`( `_Across_` )

_expression_ A variable that represents a [Range](excel.range-graph-property.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Across_|Optional| **Variant**| **True** to merge cells in each row of the specified range as separate merged cells. The default value is **False**.|

## Remarks

The value of a merged range is specified in the cell of the range's upper-left corner.


## See also


[Range Object](Excel.Range(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]