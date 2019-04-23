---
title: Color property (Excel Graph)
keywords: vbagr10.chm3077003
f1_keywords:
- vbagr10.chm3077003
ms.prod: excel
api_name:
- Excel.Color
ms.assetid: ef81e12e-1cf7-4935-e2ea-975cc8252d53
ms.date: 04/10/2019
localization_priority: Normal
---


# Color property (Excel Graph)

Returns or sets the primary color of the **Border** object, **Font** object, or the **Interior** object. Use the RGB function to create a color value. Read/write **Variant**.

## Syntax

_expression_.**Color**

_expression_ Required. An expression that returns an object in the **Applies To** list.


## Example

This example sets the color of the tick-mark labels on the value axis.

```vb
myChart.Axes(xlValue).TickLabels.Font.Color = RGB(0, 255, 0)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]