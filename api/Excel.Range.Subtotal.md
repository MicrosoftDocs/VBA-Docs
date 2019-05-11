---
title: Range.Subtotal method (Excel)
keywords: vbaxl10.chm144206
f1_keywords:
- vbaxl10.chm144206
ms.prod: excel
api_name:
- Excel.Range.Subtotal
ms.assetid: b4b7b640-5a6c-8c94-d9ab-c9a557190829
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.Subtotal method (Excel)

Creates subtotals for the range (or the current region, if the range is a single cell).


## Syntax

_expression_.**Subtotal** (_GroupBy_, _Function_, _TotalList_, _Replace_, _PageBreaks_, _SummaryBelowData_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _GroupBy_|Required| **Long**|The field to group by, as a one-based integer offset. For more information, see the example.|
| _Function_|Required| **[XlConsolidationFunction](Excel.XlConsolidationFunction.md)**| The subtotal function.|
| _TotalList_|Required| **Variant**|An array of 1-based field offsets, indicating the fields to which the subtotals are added. For more information, see the example.|
| _Replace_|Optional| **Variant**| **True** to replace existing subtotals. The default value is **True**.|
| _PageBreaks_|Optional| **Variant**| **True** to add page breaks after each group. The default value is **False**.|
| _SummaryBelowData_|Optional| **[XlSummaryRow](Excel.XlSummaryRow.md)**| Places the summary data relative to the subtotal.|

## Return value

Variant


## Example

This example creates subtotals for the selection on Sheet1. The subtotals are sums grouped by each change in field one, with the subtotals added to fields two and three.

```vb
Worksheets("Sheet1").Activate 
Selection.Subtotal GroupBy:=1, Function:=xlSum, _ 
 TotalList:=Array(2, 3)
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
