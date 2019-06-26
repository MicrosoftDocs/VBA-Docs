---
title: Legend object (PowerPoint)
keywords: vbapp10.chm709000
f1_keywords:
- vbapp10.chm709000
ms.prod: powerpoint
api_name:
- PowerPoint.Legend
ms.assetid: 7be25694-8694-049a-c31f-533fe6fd0562
ms.date: 06/08/2017
localization_priority: Normal
---


# Legend object (PowerPoint)

Represents the legend in a chart. Each chart can have only one legend.


## Remarks

 The **Legend** object contains one or more **[LegendEntry](PowerPoint.LegendEntry.md)** objects; each **LegendEntry** object contains a **[LegendKey](PowerPoint.LegendKey.md)** object.

The chart legend is not visible unless the  **[HasLegend](PowerPoint.Chart.HasLegend.md)** property is **True**. If this property is **False**, properties and methods of the **Legend** object will fail.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[Legend](PowerPoint.Chart.Legend.md)** property to return the **Legend** object. The following example sets the font style for the legend of the first chart in the active document to bold.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Legend.Font.Bold = True

    End If

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]