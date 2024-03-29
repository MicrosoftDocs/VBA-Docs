---
title: Top property (Excel Graph)
keywords: vbagr10.chm65662
f1_keywords:
- vbagr10.chm65662
api_name:
- Excel.Top
ms.assetid: 57938f4c-cd1f-b420-154d-fe4a8775c826
ms.date: 04/12/2019
ms.localizationpriority: medium
---


# Top property (Excel Graph)

For the **Application** object, the distance from the top edge of the screen to the top edge of the main Graph window. Returns or sets the position of the **Application** object. In Windows, if the application window is minimized, this property controls the position of the window icon (anywhere on the screen). Read/write **Double**.

For the **AxisTitle**, **ChartArea**, **ChartTitle**, **DataLabel**, **DataSheet**, **DisplayUnitLabel**, **Legend**, and **PlotArea** objects, the distance from the top edge of the object to the top of row 1 (on a datasheet) or the top of the chart area (on a chart). Read/write **Double**.

For the **Axis**, **LegendEntry**, and **LegendKey** objects, the distance from the top edge of the object to the top of row 1 (on a datasheet) or the top of the chart area (on a chart). Read-only **Double**.

For the **Chart** object, the distance from the top edge of the object to the top of row 1 (on a datasheet) or the top of the chart area (on a chart). Read/write **Variant**.

## Syntax

_expression_.**Top**

_expression_ Required. An expression that returns one of the above objects.

## Example

This example sets the position of the top of the chart title.

```vb
myChart.ChartTitle.Top = 10
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]