---
title: PlotArea.InsideWidth property (PowerPoint)
keywords: vbapp10.chm67205
f1_keywords:
- vbapp10.chm67205
ms.prod: powerpoint
api_name:
- PowerPoint.PlotArea.InsideWidth
ms.assetid: 99136fb4-4ee9-55e8-3c3b-bf03b95188d1
ms.date: 06/08/2017
localization_priority: Normal
---


# PlotArea.InsideWidth property (PowerPoint)

Returns or sets the inside width, in [points](../language/glossary/vbe-glossary.md#point), of the plot area. Read/write  **Double**.


## Syntax

_expression_.**InsideWidth**

_expression_ A variable that represents a '[PlotArea](PowerPoint.PlotArea.md)' object.


## Remarks

The plot area used for this measurement does not include the axis labels. The **[Width](PowerPoint.PlotArea.Width.md)** property for the plot area uses the bounding rectangle that includes the axis labels.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

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


[PlotArea Object](PowerPoint.PlotArea.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]