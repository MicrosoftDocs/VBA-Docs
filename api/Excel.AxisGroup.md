---
title: AxisGroup property (Excel Graph)
keywords: vbagr10.chm65583
f1_keywords:
- vbagr10.chm65583
ms.prod: excel
api_name:
- Excel.AxisGroup
ms.assetid: 453bc2f6-ca27-1b7c-8dc4-8a902c9445be
ms.date: 04/09/2019
localization_priority: Normal
---


# AxisGroup property (Excel Graph)

The **AxisGroup** property as it applies to the **ChartGroup**, **Series**, and **Axis** objects.

## ChartGroup and Series objects

Returns the group for the specified chart group or series. Read/write **[XlAxisGroup](excel.xlaxisgroup.md)**.

### Syntax

_expression_.**AxisGroup**

_expression_ Required. An expression that returns a **[ChartGroup](excel.chartgroup-graph-object.md)** or **[Series](excel.series-graph-object.md)** object.


## Axis object

Returns the group for the specified axis. Read-only **[XlAxisGroup](excel.xlaxisgroup.md)**.

### Syntax

_expression_.**AxisGroup**

_expression_ Required. An expression that returns an **[Axis](excel.axis-graph-object.md)** object.

### Remarks

For 3D charts, only **xlPrimary** is valid.


### Example

This example deletes the value axis if it's in the secondary group.

```vb
With myChart.Axes(xlValue) 
 If .AxisGroup = xlSecondary Then .Delete 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]