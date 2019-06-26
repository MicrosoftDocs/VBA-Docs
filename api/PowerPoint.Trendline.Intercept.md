---
title: Trendline.Intercept property (PowerPoint)
keywords: vbapp10.chm65722
f1_keywords:
- vbapp10.chm65722
ms.prod: powerpoint
api_name:
- PowerPoint.Trendline.Intercept
ms.assetid: 4ffb60a6-a5b8-9b6d-1adc-42eb6c2a7eef
ms.date: 06/08/2017
localization_priority: Normal
---


# Trendline.Intercept property (PowerPoint)

Returns or sets the point where the trendline crosses the value axis. Read/write  **Double**.


## Syntax

_expression_. `Intercept`

_expression_ A variable that represents a '[Trendline](PowerPoint.Trendline.md)' object.


## Remarks

Setting this property sets the  **[InterceptIsAuto](PowerPoint.Trendline.InterceptIsAuto.md)** property to **False**.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets trendline one for the first chart in the active document to cross the value axis at 5. You should run the example on a 2D column chart that contains a single series that has a trendline.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Trendlines(1).Intercept = 5

    End If

End With
```


## See also


[Trendline Object](PowerPoint.Trendline.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]