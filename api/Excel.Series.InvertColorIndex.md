---
title: Series.InvertColorIndex property (Excel)
keywords: vbaxl10.chm578127
f1_keywords:
- vbaxl10.chm578127
ms.prod: excel
api_name:
- Excel.Series.InvertColorIndex
ms.assetid: fa2e87a4-57ad-395d-b631-fbca99560dae
ms.date: 05/11/2019
localization_priority: Normal
---


# Series.InvertColorIndex property (Excel)

Returns or sets the fill color for negative data points in a series. Read/write.


## Syntax

_expression_.**InvertColorIndex**

_expression_ A variable that represents a **[Series](Excel.Series(object).md)** object.


## Return value

**Integer**


## Remarks

The **InvertColorIndex** property enables you to set the fill color for negative data points as a color index value from 0 to 56. For more information about color index values, see [Adding color to Excel 2007 worksheets by using the ColorIndex property](https://docs.microsoft.com/previous-versions/office/developer/office-2007/cc296089(v=office.12)). 

Instead of using the **InvertColorIndex** property, you can use the **[InvertColor](Excel.Series.InvertColor.md)** property, which enables you to set the color as a specific numeric, hexadecimal, octal, or RGB color value.

For the **InvertColorIndex** property to have an effect, the **[InvertIfNegative](Excel.Series.InvertIfNegative.md)** property of the **Series** object must also be set to **True**.


## Example

The following code example sets the fill color of negative data points in the first series of Chart 2 to magenta.

```vb
ActiveSheet.ChartObjects("Chart 2").Activate 
ActiveChart.SeriesCollection(1).InvertIfNegative = True 
ActiveChart.SeriesCollection(1).InvertColorIndex = 7
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]