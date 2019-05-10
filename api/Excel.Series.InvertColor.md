---
title: Series.InvertColor property (Excel)
keywords: vbaxl10.chm578126
f1_keywords:
- vbaxl10.chm578126
ms.prod: excel
api_name:
- Excel.Series.InvertColor
ms.assetid: 889cef2a-8211-c1b2-0668-8e0c48a894ec
ms.date: 05/11/2019
localization_priority: Normal
---


# Series.InvertColor property (Excel)

Returns or sets the fill color for negative data points in a series. Read/write.


## Syntax

_expression_.**InvertColor**

_expression_ A variable that represents a **[Series](Excel.Series(object).md)** object.


## Return value

**Integer**


## Remarks

The **InvertColor** property enables you to set the fill color for negative data points as a specific numeric, hexadecimal, octal, or RGB color value. To set the value as an RBG value, use the Visual Basic **[RGB](../language/reference/User-Interface-Help/rgb-function.md)** function. 

Instead of using the **InvertColor** property, you can use the **[InvertColorIndex](Excel.Series.InvertColorIndex.md)** property, which uses a simplier set of integer values from the current color palette.

For the **InvertColor** property to have an effect, the **[InvertIfNegative](Excel.Series.InvertIfNegative.md)** property of the **Series** object must also be set to **True**.


## Example

The following code example sets the fill color of negative data points in the first series of Chart 2 to magenta.

```vb
ActiveSheet.ChartObjects("Chart 2").Activate 
ActiveChart.SeriesCollection(1).InvertIfNegative = True 
ActiveChart.SeriesCollection(1).InvertColor = RGB(255, 0, 255)
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]