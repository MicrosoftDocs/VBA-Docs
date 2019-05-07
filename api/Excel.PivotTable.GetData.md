---
title: PivotTable.GetData method (Excel)
keywords: vbaxl10.chm235110
f1_keywords:
- vbaxl10.chm235110
ms.prod: excel
api_name:
- Excel.PivotTable.GetData
ms.assetid: c3b88918-c515-a976-5f2e-107b981ac76f
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.GetData method (Excel)

Returns the value for the data filed in a PivotTable.


## Syntax

_expression_.**GetData** (_Name_)

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|Describes a single cell in the PivotTable report, using syntax similar to the **[PivotSelect](Excel.PivotTable.PivotSelect.md)** method or the PivotTable report references in calculated item formulas.|

## Return value

Double


## Example

This example shows the sum of revenues for apples in January (Data field = Revenue, Product = Apples, Month = January).

```vb
Msgbox ActiveSheet.PivotTables(1) _ 
 .GetData("'Sum of Revenue' Apples January")
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]