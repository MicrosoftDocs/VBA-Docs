---
title: Borders.Color property (Excel)
keywords: vbaxl10.chm181073
f1_keywords:
- vbaxl10.chm181073
api_name:
- Excel.Borders.Color
ms.assetid: 3ee1bce3-56e2-c93f-432f-8f1d037a7624
ms.date: 04/13/2019
ms.localizationpriority: medium
---


# Borders.Color property (Excel)

Returns or sets the primary color of the object, as shown in the table in the remarks section. Use the **[RGB](../Language/Reference/User-Interface-Help/rgb-function.md)** function to create a color value. Read/write **Variant**.


## Syntax

_expression_.**Color**

_expression_ An expression that returns a **[Borders](Excel.Borders.md)** object.


## Remarks

|Object|Color|
|:-----|:-----|
| **Border**|The color of the border.|
| **Borders**|The color of all four borders of a range. If they're not all the same color, **Color** returns 0 (zero).|
| **Font**|The color of the font.|
| **Interior**|The cell shading color or the drawing object fill color.|
| **Tab**|The color of the tab.|

## Example

This example sets the color of the tick-mark labels on the value axis on Chart1.

```vb
Charts("Chart1").Axes(xlValue).TickLabels.Font.Color = _ 
 RGB(0, 255, 0)
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]