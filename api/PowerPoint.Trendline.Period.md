---
title: Trendline.Period property (PowerPoint)
keywords: vbapp10.chm65720
f1_keywords:
- vbapp10.chm65720
ms.prod: powerpoint
api_name:
- PowerPoint.Trendline.Period
ms.assetid: 7482f5c1-412f-8653-b5f3-1b672125b3e5
ms.date: 06/08/2017
localization_priority: Normal
---


# Trendline.Period property (PowerPoint)

Returns or sets the period for the moving-average trendline. Read/write  **Long**.


## Syntax

_expression_. `Period`

_expression_ A variable that represents a '[Trendline](PowerPoint.Trendline.md)' object.


## Remarks

This property can have a value from 2 through 255. 


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the period for the moving-average trendline on the first chart in the active document. You should run the example on a 2D column chart that has a single series that contains 10 data points and a moving-average trendline.




```vb
With ActiveDocument.InlineShapes(1) 
    If .HasChart Then 
        With .Chart.SeriesCollection(1).Trendlines(1) 
            If .Type = xlMovingAvg Then .Period = 5 
        End With 
    End If 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]