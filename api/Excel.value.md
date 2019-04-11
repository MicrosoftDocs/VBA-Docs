---
title: Value property (Excel Graph)
keywords: vbagr10.chm65542
f1_keywords:
- vbagr10.chm65542
ms.prod: excel
ms.assetid: c88258bc-7088-71df-87e7-49239239de76
ms.date: 04/12/2019
localization_priority: Normal
---


# Value property (Excel Graph)

Returns the value of the specified cell. If the cell is empty, **Value** returns the value Empty (use the **IsEmpty** function to test for this case). 

If the **Range** object contains more than one cell, this property returns an array of values (use the **IsArray** function to test for this case). Read/write **Variant**.

## Syntax

_expression_.**Value** (_RangeValueDataType_)

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Parameters

|Name|Required/Optional|Data type|
|:---|:----------------|:--------|
|_RangeValueDataType_ | Optional |**Variant** |

## Example

This example sets the value of cell A1 on the datasheet to 3.14159.

```vb
myChart.Application.DataSheet.Range("A1").Value = 3.14159
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]