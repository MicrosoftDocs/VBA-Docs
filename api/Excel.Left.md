---
title: Left property (Excel Graph)
keywords: vbagr10.chm65663
f1_keywords:
- vbagr10.chm65663
ms.prod: excel
api_name:
- Excel.Left
ms.assetid: 9d300adc-3d72-02d5-e39c-c40e21b7e8d5
ms.date: 04/11/2019
localization_priority: Normal
---


# Left property (Excel Graph)

The **Left** property as it applies to the following objects.

## Application and Datasheet objects

Returns or sets the distance from the left edge of the screen to the left edge of the main Graph window. Read/write **Double**.

### Syntax

_expression_.**Left**

_expression_ An expression that returns the **[Application](excel.application-graph-object.md)** or **[DataSheet](excel.datasheet-graph-object.md)** object.

### Remarks

If the window is maximized, **Application.Left** returns a negative number that varies based on the width of the window border. Setting **Application.Left** to 0 (zero) will make the window a tiny bit smaller than it would be if the application window were maximized. In other words, if **Application.Left** is 0, the left border of the main Graph window will just barely be visible on the screen.

If the Graph window is minimized, **Application.Left** controls the position of the window icon.


## AxisTitle, ChartArea, ChartTitle, DataLabel, DisplayUnitLabel, Legend, and PlotArea objects

Returns or sets the distance from the left edge of the object to the left edge of the chart area. Read/write **Double**.

### Syntax

_expression_.**Left**

_expression_ Required. An expression that returns one of the above objects.

### Example

This example aligns the left edge of the chart title with the left edge of the chart area.

```vb
myChart.ChartTitle.Left = 0 

```

## Axis, LegendEntry, and LegendKey objects

Returns or sets the distance from the left edge of the object to the left edge of the chart area. Read-only **Double**.

### Syntax

_expression_.**Left**

_expression_ Required. An expression that returns one of the above objects.


## Chart object

Returns or sets the distance from the left edge of the object to the left edge of the Graph window. Read/write **Variant**.

### Syntax

_expression_.**Left**

_expression_ Required. An expression that returns a **[Chart](Excel.Chart-graph-object.md)** object.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]