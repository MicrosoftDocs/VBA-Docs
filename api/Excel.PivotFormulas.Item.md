---
title: PivotFormulas.Item method (Excel)
keywords: vbaxl10.chm233075
f1_keywords:
- vbaxl10.chm233075
ms.prod: excel
api_name:
- Excel.PivotFormulas.Item
ms.assetid: 023f5702-9e18-f5d1-82b8-2603a98eb0b2
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotFormulas.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[PivotFormulas](Excel.PivotFormulas.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

## Return value

A **[PivotFormula](Excel.PivotFormula.md)** object contained by the collection.


## Example

This example displays the first formula for PivotTable one on worksheet one.

```vb
MsgBox Worksheets(1).PivotTables(1).PivotFormulas.Item(1).Formula
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]