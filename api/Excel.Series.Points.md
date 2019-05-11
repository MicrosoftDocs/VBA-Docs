---
title: Series.Points method (Excel)
keywords: vbaxl10.chm578104
f1_keywords:
- vbaxl10.chm578104
ms.prod: excel
api_name:
- Excel.Series.Points
ms.assetid: 9b6f08a1-3fbe-e9bc-a509-345a3d2d78b3
ms.date: 05/11/2019
localization_priority: Normal
---


# Series.Points method (Excel)

Returns an object that represents a single point (a **[Point](Excel.Point(object).md)** object) or a collection of all the points (a **[Points](Excel.Points(object).md)** collection) in the series. Read-only.


## Syntax

_expression_.**Points** (_Index_)

 _expression_ An expression that returns a **[Series](Excel.Series(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the point.|

## Return value

Object


## Example

This example applies a data label to point one in series one on Chart1.

```vb
Charts("Chart1").SeriesCollection(1).Points(1).ApplyDataLabels
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]