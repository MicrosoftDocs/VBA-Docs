---
title: Border.ColorIndex property (Excel)
keywords: vbaxl10.chm547074
f1_keywords:
- vbaxl10.chm547074
ms.prod: excel
api_name:
- Excel.Border.ColorIndex
ms.assetid: 35e2dbbf-fd35-a08c-a969-bd08d0544d97
ms.date: 03/07/2019
localization_priority: Normal
---


# Border.ColorIndex property (Excel)

Returns or sets a **Variant** value that represents the color of the border.


## Syntax

_expression_.**ColorIndex**

_expression_ A variable that represents a **[Border](Excel.Border(object).md)** object.


## Remarks

The color is specified as an index value into the current color palette, or as one of the following **[XlColorIndex](Excel.XlColorIndex.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone**.
    
> [!IMPORTANT] 
> Note that the visual properties of a **Border** object are interlocked; that is, changing one property can induce changes in another. In most cases, the induced changes serve to make the border visible (which may or may not be desirable). However, other (more unexpected) results are possible. For an example, see the **[Border](excel.border(object).md)** object.

## Example

This example sets the color of the major gridlines for the value axis on Chart1.

```vb
With Charts("Chart1").Axes(xlValue) 
 If .HasMajorGridlines Then 
 .MajorGridlines.Border.ColorIndex = 5 'set color to blue 
 End If 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]