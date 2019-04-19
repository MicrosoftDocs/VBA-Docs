---
title: Chart.ChartGroups method (Excel)
keywords: vbaxl10.chm149087
f1_keywords:
- vbaxl10.chm149087
ms.prod: excel
api_name:
- Excel.Chart.ChartGroups
ms.assetid: dffa4fc3-b2db-eb50-b309-95e99972525f
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.ChartGroups method (Excel)

Returns an object that represents either a single chart group (a **[ChartGroup](Excel.ChartGroup(object).md)** object) or a collection of all the chart groups in the chart (a **[ChartGroups](Excel.ChartGroups(object).md)** object). The returned collection includes every type of group.


## Syntax

_expression_.**ChartGroups** (_Index_)

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The chart group number.|


## Return value

Object


## Example

This example turns on up and down bars for chart group one on Chart1 and then sets their colors. The example should be run on a 2D line chart containing two series that intersect at one or more data points.

```vb
With Charts("Chart1").ChartGroups(1) 
 .HasUpDownBars = True 
 .DownBars.Interior.ColorIndex = 3 
 .UpBars.Interior.ColorIndex = 5 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]