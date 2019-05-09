---
title: Border.Color property (Excel)
keywords: vbaxl10.chm547073
f1_keywords:
- vbaxl10.chm547073
ms.prod: excel
api_name:
- Excel.Border.Color
ms.assetid: ca90fc42-2a7a-d43e-9c2c-0055f6bf9010
ms.date: 04/13/2019
localization_priority: Normal
---


# Border.Color property (Excel)

Returns or sets the primary color of the object, as shown in the table in the remarks section. Use the **[RGB](../Language/Reference/User-Interface-Help/rgb-function.md)** function to create a color value. Read/write **Variant**.


## Syntax

_expression_.**Color**

_expression_ An expression that returns a **[Border](Excel.Border(object).md)** object.


## Remarks

|Object|Color|
|:-----|:-----|
| **Border**|The color of the border.|
| **Borders**|The color of all four borders of a range. If they're not all the same color, **Color** returns 0 (zero).|
| **Font**|The color of the font.|
| **Interior**|The cell shading color or the drawing object fill color.|
| **Tab**|The color of the tab.|

> [!IMPORTANT] 
> Note that the visual properties of a **Border** object are interlocked; that is, changing one property can induce changes in another. In most cases, the induced changes serve to make the border visible (which may or may not be desirable). However, other (more unexpected) results are possible. For an example, see the **[Border](excel.border(object).md)** object.


## Example

This example sets the color of the tick-mark labels on the value axis on Chart1.

```vb
Charts("Chart1").Axes(xlValue).TickLabels.Font.Color = _ 
 RGB(0, 255, 0)
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
