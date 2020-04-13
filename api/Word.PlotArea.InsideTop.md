---
title: PlotArea.InsideTop property (Word)
keywords: vbawd10.chm53479044
f1_keywords:
- vbawd10.chm53479044
ms.prod: word
api_name:
- Word.PlotArea.InsideTop
ms.assetid: 803b9238-b076-807f-7c27-5df6fcce878c
ms.date: 06/08/2017
localization_priority: Normal
---


# PlotArea.InsideTop property (Word)

Returns or sets the distance, in [points](../language/glossary/vbe-glossary.md#point), from the chart edge to the inside top edge of the plot area. Read/write  **Double**.


## Syntax

_expression_.**InsideTop**

_expression_ A variable that represents a '[PlotArea](Word.PlotArea.md)' object.


## Remarks

The plot area used for this measurement does not include the axis labels. The **[Top](Word.PlotArea.Top.md)** property for the plot area uses the bounding rectangle that includes the axis labels.


## Example

The following example draws a dotted rectangle around the inside of the plot area for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart 
 Set pa = .PlotArea 
 With .Shapes.AddShape(msoShapeRectangle, _ 
 pa.InsideLeft, pa.InsideTop, _ 
 pa.InsideWidth, pa.InsideHeight) 
 .Fill.Transparency = 1 
 .Line.DashStyle = msoLineDashDot 
 End With 
 End With 
 End If 
End With
```


## See also


[PlotArea Object](Word.PlotArea.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]