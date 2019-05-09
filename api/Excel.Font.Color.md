---
title: Font.Color property (Excel)
keywords: vbaxl10.chm559075
f1_keywords:
- vbaxl10.chm559075
ms.prod: excel
api_name:
- Excel.Font.Color
ms.assetid: a6acd8b8-f04b-6d43-15d4-78ee20b0b14d
ms.date: 04/26/2019
localization_priority: Normal
---


# Font.Color property (Excel)

Returns or sets the primary color of the object, as shown in the table in the remarks section. Use the **[RGB](../Language/Reference/User-Interface-Help/rgb-function.md)** function to create a color value. Read/write **Variant**.


## Syntax

_expression_.**Color**

_expression_ An expression that returns a **[Font](excel.font(object).md)** object.


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
