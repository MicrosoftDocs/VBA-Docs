---
title: PlotArea.InsideHeight property (Excel)
keywords: vbaxl10.chm618091
f1_keywords:
- vbaxl10.chm618091
ms.prod: excel
api_name:
- Excel.PlotArea.InsideHeight
ms.assetid: a9b2e591-afc5-331e-86b5-bbeb47696c3d
ms.date: 05/09/2019
localization_priority: Normal
---


# PlotArea.InsideHeight property (Excel)

Returns the inside height of the plot area, in [points](../language/glossary/vbe-glossary.md#point). Read/write **Double**.


## Syntax

_expression_.**InsideHeight**

_expression_ A variable that represents a **[PlotArea](Excel.PlotArea(object).md)** object.


## Remarks

The plot area used for this measurement doesn't include the axis labels. The **Height** property for the plot area uses the bounding rectangle that includes the axis labels.


## Example

This example draws a dotted rectangle around the inside of the plot area on Chart1.

```vb
With Charts("chart1") 
 Set pa = .PlotArea 
 With .Shapes.AddShape(msoShapeRectangle, _ 
 pa.InsideLeft, pa.InsideTop, _ 
 pa.InsideWidth, pa.InsideHeight) 
 .Fill.Transparency = 1 
 .Line.DashStyle = msoLineDashDot 
 End With 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]