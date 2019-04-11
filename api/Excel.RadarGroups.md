---
title: RadarGroups method (Excel Graph)
keywords: vbagr10.chm3077630
f1_keywords:
- vbagr10.chm3077630
ms.prod: excel
api_name:
- Excel.RadarGroups
ms.assetid: 5fbca532-ae99-fb41-dd1d-2d24909bf073
ms.date: 04/09/2019
localization_priority: Normal
---


# RadarGroups method (Excel Graph)

On a 2D chart, returns an object that represents either a single radar chart group or a collection of the radar chart groups.

## Syntax

_expression_.**RadarGroups** (_Index_)

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_ | Optional | **Variant**| Specifies the chart group.| 

## Example

This example sets radar group one to use a different color for each data marker. The example should be run on a 2D chart.

```vb
myChart.RadarGroups(1).VaryByCategories = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]