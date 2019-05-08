---
title: PlotArea.InsideWidth property (Word)
keywords: vbawd10.chm53479045
f1_keywords:
- vbawd10.chm53479045
ms.prod: word
api_name:
- Word.PlotArea.InsideWidth
ms.assetid: acdb721d-73a9-15af-d833-d044e83b3c87
ms.date: 06/08/2017
localization_priority: Normal
---


# PlotArea.InsideWidth property (Word)

Returns or sets the inside width, in [points](../language/glossary/vbe-glossary.md#point), of the plot area. Read/write  **Double**.


## Syntax

_expression_.**InsideWidth**

_expression_ A variable that represents a '[PlotArea](Word.PlotArea.md)' object.


## Remarks

The plot area used for this measurement does not include the axis labels. The  **[Width](Word.PlotArea.Width.md)** property for the plot area uses the bounding rectangle that includes the axis labels.


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