---
title: LegendEntries method (Excel Graph)
keywords: vbagr10.chm3077622
f1_keywords:
- vbagr10.chm3077622
ms.prod: excel
api_name:
- Excel.LegendEntries
ms.assetid: 6419aa89-6e59-dc04-ab79-67a0a38cad6f
ms.date: 04/09/2019
localization_priority: Normal
---


# LegendEntries method (Excel Graph)

Returns an object that represents either a single legend entry or a collection of legend entries for the legend.

## Syntax

_expression_.**LegendEntries** (_Index_)

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_ | Optional |**Variant**|The number of the legend entry.|

## Example

This example sets the font for legend entry one.

```vb
myChart.Legend.LegendEntries(1).Font.Name = "Arial"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]