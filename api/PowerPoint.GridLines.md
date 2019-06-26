---
title: Gridlines object (PowerPoint)
keywords: vbapp10.chm705000
f1_keywords:
- vbapp10.chm705000
ms.prod: powerpoint
api_name:
- PowerPoint.GridLines
ms.assetid: 10b45c4c-05a3-f722-15ca-ad0242625edb
ms.date: 06/08/2017
localization_priority: Normal
---


# Gridlines object (PowerPoint)

Represents major or minor gridlines on a chart axis.


## Remarks

 Gridlines extend the tick marks on a chart axis to make it easier to see the values associated with the data markers. This object is not a collection. There is no object that represents a single gridline; you either enable all gridlines for an axis or disable all of them.

Use the  **[MajorGridlines](PowerPoint.Axis.MajorGridlines.md)** property to return the **GridLines** object that represents the major gridlines for the axis. Use the **[MinorGridlines](PowerPoint.Axis.MinorGridlines.md)** property to return the **GridLines** object that represents the minor gridlines. It is possible to return both major and minor gridlines at the same time.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables major gridlines for the category axis of the first chart in the active document and then formats the gridlines to be blue dashed lines.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlCategory)

            .HasMajorGridlines = True

            .MajorGridlines.Border.Color = RGB(0, 0, 255)

            .MajorGridlines.Border.LineStyle = xlDash

        End With

    End If

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]